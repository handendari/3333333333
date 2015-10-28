VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_employee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER KARYAWAN"
   ClientHeight    =   10815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14850
   Icon            =   "frmMstEmployee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10815
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   Begin prj_panji.vbButton cmdExit 
      Height          =   705
      Left            =   13770
      TabIndex        =   139
      Top             =   10080
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
      MICON           =   "frmMstEmployee.frx":058A
      PICN            =   "frmMstEmployee.frx":05A6
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
      Height          =   9285
      Left            =   180
      TabIndex        =   47
      Top             =   720
      Width           =   14565
      _ExtentX        =   25691
      _ExtentY        =   16378
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      TabCaption(0)   =   "KARYAWAN"
      TabPicture(0)   =   "frmMstEmployee.frx":1638
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_employee"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TDBGrid_Employee"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_entry_employee"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TDBCombo_company"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txt_company_name"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra_status_emp"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CommonDialog1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "KELUARGA"
      TabPicture(1)   =   "frmMstEmployee.frx":1654
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_fam"
      Tab(1).Control(1)=   "TDBGrid_Family"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(3)=   "fra_entry_family"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "PENDIDIKAN"
      TabPicture(2)   =   "frmMstEmployee.frx":1670
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5(0)"
      Tab(2).Control(1)=   "fra_entry_education"
      Tab(2).Control(2)=   "TDBGrid_Education"
      Tab(2).Control(3)=   "lbl_edu"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "KEAHLIAN"
      TabPicture(3)   =   "frmMstEmployee.frx":168C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fra_entry_skill"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).Control(2)=   "TDBGrid_Skill"
      Tab(3).Control(3)=   "lbl_skill"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "PENGALAMAN KERJA"
      TabPicture(4)   =   "frmMstEmployee.frx":16A8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fra_entry_job"
      Tab(4).Control(1)=   "Frame5(1)"
      Tab(4).Control(2)=   "TDBGrid_Job"
      Tab(4).Control(3)=   "lbl_exp"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "JABATAN"
      TabPicture(5)   =   "frmMstEmployee.frx":16C4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame5(5)"
      Tab(5).Control(1)=   "fra_entry_title"
      Tab(5).Control(2)=   "TDBGrid_Title"
      Tab(5).Control(3)=   "lbl_jbtn"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "EMP. GRADE"
      TabPicture(6)   =   "frmMstEmployee.frx":16E0
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame5(4)"
      Tab(6).Control(1)=   "fra_entry_grade"
      Tab(6).Control(2)=   "TDBGrid_Grade"
      Tab(6).Control(3)=   "lbl_grade"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "TRAINING"
      TabPicture(7)   =   "frmMstEmployee.frx":16FC
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fra_entry_training"
      Tab(7).Control(1)=   "Frame5(2)"
      Tab(7).Control(2)=   "TDBGrid_Training"
      Tab(7).Control(3)=   "lbl_training"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   "KONTRAK"
      TabPicture(8)   =   "frmMstEmployee.frx":1718
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lbl_kontrak"
      Tab(8).Control(1)=   "TDBGrid_Contract"
      Tab(8).Control(2)=   "fra_entry_contract"
      Tab(8).Control(3)=   "Frame5(3)"
      Tab(8).ControlCount=   4
      Begin VB.Frame Frame5 
         Caption         =   "Data Control Button"
         Height          =   1335
         Index           =   5
         Left            =   -74760
         TabIndex        =   254
         Top             =   7650
         Width           =   13995
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Index           =   8
            Left            =   540
            TabIndex        =   255
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
            MICON           =   "frmMstEmployee.frx":1734
            PICN            =   "frmMstEmployee.frx":1750
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
            Index           =   8
            Left            =   1560
            TabIndex        =   256
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
            MICON           =   "frmMstEmployee.frx":27E2
            PICN            =   "frmMstEmployee.frx":27FE
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
            Index           =   8
            Left            =   2580
            TabIndex        =   257
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
            MICON           =   "frmMstEmployee.frx":3890
            PICN            =   "frmMstEmployee.frx":38AC
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
            Index           =   8
            Left            =   3600
            TabIndex        =   258
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
            MICON           =   "frmMstEmployee.frx":493E
            PICN            =   "frmMstEmployee.frx":495A
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
            Index           =   8
            Left            =   4620
            TabIndex        =   259
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
            MICON           =   "frmMstEmployee.frx":59EC
            PICN            =   "frmMstEmployee.frx":5A08
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
      Begin VB.Frame fra_entry_title 
         Height          =   2355
         Left            =   -74760
         TabIndex        =   246
         Top             =   5220
         Width           =   13965
         Begin VB.TextBox txt_title 
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
            Height          =   405
            Left            =   5820
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   247
            Top             =   1080
            Width           =   3705
         End
         Begin VB.TextBox txt_description_title 
            Appearance      =   0  'Flat
            Height          =   555
            Left            =   5820
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   251
            Top             =   1560
            Width           =   3705
         End
         Begin MSComCtl2.DTPicker DTPicker_title 
            Height          =   315
            Left            =   5820
            TabIndex        =   248
            Top             =   360
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_title_emp 
            Height          =   375
            Left            =   5820
            OleObjectBlob   =   "frmMstEmployee.frx":6A9A
            TabIndex        =   249
            Top             =   720
            Width           =   1545
         End
         Begin VB.Label Label63 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JABATAN*"
            Height          =   195
            Left            =   4185
            TabIndex        =   253
            Top             =   750
            Width           =   1110
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TANGGAL*"
            Height          =   195
            Left            =   4470
            TabIndex        =   252
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label61 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   3825
            TabIndex        =   250
            Top             =   1590
            Width           =   1470
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Data Control Button"
         Height          =   1335
         Index           =   4
         Left            =   -74700
         TabIndex        =   239
         Top             =   7590
         Width           =   13995
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Index           =   5
            Left            =   540
            TabIndex        =   240
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
            MICON           =   "frmMstEmployee.frx":89FA
            PICN            =   "frmMstEmployee.frx":8A16
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
            Index           =   5
            Left            =   1560
            TabIndex        =   241
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
            MICON           =   "frmMstEmployee.frx":9AA8
            PICN            =   "frmMstEmployee.frx":9AC4
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
            Index           =   5
            Left            =   2580
            TabIndex        =   242
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
            MICON           =   "frmMstEmployee.frx":AB56
            PICN            =   "frmMstEmployee.frx":AB72
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
            Index           =   5
            Left            =   3600
            TabIndex        =   243
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
            MICON           =   "frmMstEmployee.frx":BC04
            PICN            =   "frmMstEmployee.frx":BC20
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
            Index           =   5
            Left            =   4620
            TabIndex        =   244
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
            MICON           =   "frmMstEmployee.frx":CCB2
            PICN            =   "frmMstEmployee.frx":CCCE
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
      Begin VB.Frame fra_entry_grade 
         Height          =   2355
         Left            =   -74700
         TabIndex        =   231
         Top             =   5160
         Width           =   13965
         Begin VB.TextBox txt_grade_description 
            Appearance      =   0  'Flat
            Height          =   555
            Left            =   5820
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   233
            Top             =   1530
            Width           =   3705
         End
         Begin VB.TextBox txt_grade_name1 
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
            Height          =   405
            Left            =   5820
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   232
            Top             =   1080
            Width           =   3705
         End
         Begin MSComCtl2.DTPicker DTPicker_grade 
            Height          =   315
            Left            =   5820
            TabIndex        =   234
            Top             =   360
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_grade1 
            Height          =   375
            Left            =   5820
            OleObjectBlob   =   "frmMstEmployee.frx":DD60
            TabIndex        =   235
            Top             =   720
            Width           =   1545
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   3825
            TabIndex        =   238
            Top             =   1590
            Width           =   1110
         End
         Begin VB.Label Label60 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TANGGAL*"
            Height          =   195
            Left            =   4050
            TabIndex        =   237
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "GRADE*"
            Height          =   195
            Left            =   4260
            TabIndex        =   236
            Top             =   750
            Width           =   630
         End
      End
      Begin VB.Frame fra_entry_training 
         Height          =   2685
         Left            =   -74730
         TabIndex        =   217
         Top             =   4830
         Width           =   13965
         Begin VB.TextBox txt_training_place 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5820
            TabIndex        =   221
            Top             =   1800
            Width           =   3705
         End
         Begin VB.TextBox txt_training_organize 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            TabIndex        =   220
            Top             =   1440
            Width           =   3705
         End
         Begin VB.TextBox txt_training_subject 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            TabIndex        =   219
            Top             =   1080
            Width           =   3705
         End
         Begin VB.TextBox txt_training_value 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5820
            TabIndex        =   218
            Top             =   2130
            Width           =   3705
         End
         Begin MSComCtl2.DTPicker DTPicker_training_start 
            Height          =   315
            Left            =   5820
            TabIndex        =   222
            Top             =   360
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_training_end 
            Height          =   315
            Left            =   5820
            TabIndex        =   223
            Top             =   720
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TEMPAT TRAINING*"
            Height          =   195
            Left            =   4110
            TabIndex        =   229
            Top             =   1800
            Width           =   1545
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PENYELENGGARA*"
            Height          =   195
            Left            =   4170
            TabIndex        =   228
            Top             =   1440
            Width           =   1485
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL. MULAI*"
            Height          =   195
            Left            =   4680
            TabIndex        =   227
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA TRAINING*"
            Height          =   195
            Left            =   4305
            TabIndex        =   226
            Top             =   1080
            Width           =   1350
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL. BERAKHIR*"
            Height          =   195
            Left            =   4365
            TabIndex        =   225
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NILAI"
            Height          =   195
            Left            =   5280
            TabIndex        =   224
            Top             =   2130
            Width           =   405
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Data Control Button"
         Height          =   1335
         Index           =   2
         Left            =   -74730
         TabIndex        =   211
         Top             =   7560
         Width           =   13995
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Index           =   6
            Left            =   540
            TabIndex        =   212
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
            MICON           =   "frmMstEmployee.frx":FCBD
            PICN            =   "frmMstEmployee.frx":FCD9
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
            Index           =   6
            Left            =   1560
            TabIndex        =   213
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
            MICON           =   "frmMstEmployee.frx":10D6B
            PICN            =   "frmMstEmployee.frx":10D87
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
            Index           =   6
            Left            =   2580
            TabIndex        =   214
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
            MICON           =   "frmMstEmployee.frx":11E19
            PICN            =   "frmMstEmployee.frx":11E35
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
            Index           =   6
            Left            =   3600
            TabIndex        =   215
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
            MICON           =   "frmMstEmployee.frx":12EC7
            PICN            =   "frmMstEmployee.frx":12EE3
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
            Index           =   6
            Left            =   4620
            TabIndex        =   216
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
            MICON           =   "frmMstEmployee.frx":13F75
            PICN            =   "frmMstEmployee.frx":13F91
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
         Index           =   3
         Left            =   -74730
         TabIndex        =   204
         Top             =   7620
         Width           =   13995
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Index           =   7
            Left            =   540
            TabIndex        =   205
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
            MICON           =   "frmMstEmployee.frx":15023
            PICN            =   "frmMstEmployee.frx":1503F
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
            Index           =   7
            Left            =   1560
            TabIndex        =   206
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
            MICON           =   "frmMstEmployee.frx":160D1
            PICN            =   "frmMstEmployee.frx":160ED
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
            Index           =   7
            Left            =   2580
            TabIndex        =   207
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
            MICON           =   "frmMstEmployee.frx":1717F
            PICN            =   "frmMstEmployee.frx":1719B
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
            Index           =   7
            Left            =   3600
            TabIndex        =   208
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
            MICON           =   "frmMstEmployee.frx":1822D
            PICN            =   "frmMstEmployee.frx":18249
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
            Index           =   7
            Left            =   4620
            TabIndex        =   209
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
            MICON           =   "frmMstEmployee.frx":192DB
            PICN            =   "frmMstEmployee.frx":192F7
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
      Begin VB.Frame fra_entry_contract 
         Height          =   2355
         Left            =   -74730
         TabIndex        =   195
         Top             =   5190
         Width           =   13965
         Begin VB.TextBox txt_contract_description 
            Appearance      =   0  'Flat
            Height          =   555
            Left            =   5820
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   197
            Top             =   1470
            Width           =   3705
         End
         Begin VB.TextBox txt_contract_no 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            TabIndex        =   196
            Top             =   1080
            Width           =   3705
         End
         Begin MSComCtl2.DTPicker DTPicker_contract_start 
            Height          =   315
            Left            =   5820
            TabIndex        =   198
            Top             =   360
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_contract_end 
            Height          =   315
            Left            =   5820
            TabIndex        =   199
            Top             =   720
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   4545
            TabIndex        =   203
            Top             =   1500
            Width           =   1110
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL. AKHIR*"
            Height          =   195
            Left            =   4665
            TabIndex        =   202
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NO. KONTRAK*"
            Height          =   195
            Left            =   4455
            TabIndex        =   201
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL. AWAL*"
            Height          =   195
            Left            =   4680
            TabIndex        =   200
            Top             =   360
            Width           =   930
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9930
         Top             =   450
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fra_entry_job 
         Height          =   4035
         Left            =   -74700
         TabIndex        =   126
         Top             =   3510
         Width           =   13965
         Begin VB.TextBox txt_job_description 
            Appearance      =   0  'Flat
            Height          =   555
            Left            =   5820
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   68
            Top             =   3150
            Width           =   3705
         End
         Begin VB.TextBox txt_job_reason 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5820
            TabIndex        =   67
            Top             =   2820
            Width           =   3705
         End
         Begin VB.TextBox txt_job_company 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            TabIndex        =   62
            Top             =   1080
            Width           =   3705
         End
         Begin VB.TextBox txt_job_title 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5820
            TabIndex        =   65
            Top             =   2160
            Width           =   3705
         End
         Begin VB.TextBox txt_job_dept 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            TabIndex        =   64
            Top             =   1800
            Width           =   3705
         End
         Begin VB.TextBox txt_job_line 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            TabIndex        =   63
            Top             =   1440
            Width           =   3705
         End
         Begin VB.TextBox txt_job_salary 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5820
            TabIndex        =   66
            Top             =   2490
            Width           =   3705
         End
         Begin MSComCtl2.DTPicker DTPicker_job_start 
            Height          =   315
            Left            =   5820
            TabIndex        =   60
            Top             =   360
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_job_end 
            Height          =   315
            Left            =   5820
            TabIndex        =   61
            Top             =   720
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   4545
            TabIndex        =   136
            Top             =   3180
            Width           =   1110
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ALASAN KELUAR"
            Height          =   195
            Left            =   4335
            TabIndex        =   135
            Top             =   2820
            Width           =   1320
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PERUSAHAAN*"
            Height          =   195
            Left            =   4485
            TabIndex        =   134
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JABATAN TERAKHIR"
            Height          =   195
            Left            =   4065
            TabIndex        =   132
            Top             =   2160
            Width           =   1590
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DEPARTEMEN"
            Height          =   195
            Left            =   4530
            TabIndex        =   131
            Top             =   1800
            Width           =   1125
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL. MASUK*"
            Height          =   195
            Left            =   4635
            TabIndex        =   130
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BIDANG USAHA"
            Height          =   195
            Left            =   4440
            TabIndex        =   129
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL. KELUAR*"
            Height          =   195
            Left            =   4545
            TabIndex        =   128
            Top             =   720
            Width           =   1110
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UPAH TERAKHIR"
            Height          =   195
            Left            =   4320
            TabIndex        =   127
            Top             =   2490
            Width           =   1320
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Data Control Button"
         Height          =   1335
         Index           =   1
         Left            =   -74700
         TabIndex        =   120
         Top             =   7620
         Width           =   13995
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Index           =   4
            Left            =   540
            TabIndex        =   121
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
            MICON           =   "frmMstEmployee.frx":1A389
            PICN            =   "frmMstEmployee.frx":1A3A5
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
            Index           =   4
            Left            =   1560
            TabIndex        =   122
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
            MICON           =   "frmMstEmployee.frx":1B437
            PICN            =   "frmMstEmployee.frx":1B453
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
            Index           =   4
            Left            =   2580
            TabIndex        =   123
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
            MICON           =   "frmMstEmployee.frx":1C4E5
            PICN            =   "frmMstEmployee.frx":1C501
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
            Index           =   4
            Left            =   3600
            TabIndex        =   124
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
            MICON           =   "frmMstEmployee.frx":1D593
            PICN            =   "frmMstEmployee.frx":1D5AF
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
            Index           =   4
            Left            =   4620
            TabIndex        =   125
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
            MICON           =   "frmMstEmployee.frx":1E641
            PICN            =   "frmMstEmployee.frx":1E65D
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
      Begin VB.Frame fra_entry_skill 
         Height          =   2325
         Left            =   -74730
         TabIndex        =   115
         Top             =   5250
         Width           =   13965
         Begin VB.TextBox txt_skill_description 
            Appearance      =   0  'Flat
            Height          =   795
            Left            =   5820
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   59
            Top             =   1200
            Width           =   3705
         End
         Begin VB.ComboBox cbo_skill_level 
            Height          =   315
            ItemData        =   "frmMstEmployee.frx":1F6EF
            Left            =   5820
            List            =   "frmMstEmployee.frx":1F6FF
            TabIndex        =   57
            Text            =   "Excellent"
            Top             =   810
            Width           =   1545
         End
         Begin VB.TextBox txt_skill_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            TabIndex        =   55
            Top             =   420
            Width           =   3705
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   4575
            TabIndex        =   119
            Top             =   1230
            Width           =   1110
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "LEVEL*"
            Height          =   195
            Left            =   4500
            TabIndex        =   118
            Top             =   810
            Width           =   1185
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KEAHLIAN*"
            Height          =   195
            Left            =   4830
            TabIndex        =   116
            Top             =   450
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74730
         TabIndex        =   109
         Top             =   7650
         Width           =   13995
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Index           =   3
            Left            =   540
            TabIndex        =   110
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
            MICON           =   "frmMstEmployee.frx":1F723
            PICN            =   "frmMstEmployee.frx":1F73F
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
            Index           =   3
            Left            =   1560
            TabIndex        =   111
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
            MICON           =   "frmMstEmployee.frx":207D1
            PICN            =   "frmMstEmployee.frx":207ED
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
            Index           =   3
            Left            =   2580
            TabIndex        =   112
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
            MICON           =   "frmMstEmployee.frx":2187F
            PICN            =   "frmMstEmployee.frx":2189B
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
            Index           =   3
            Left            =   3600
            TabIndex        =   113
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
            MICON           =   "frmMstEmployee.frx":2292D
            PICN            =   "frmMstEmployee.frx":22949
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
            Index           =   3
            Left            =   4620
            TabIndex        =   114
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
            MICON           =   "frmMstEmployee.frx":239DB
            PICN            =   "frmMstEmployee.frx":239F7
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
         Index           =   0
         Left            =   -74730
         TabIndex        =   101
         Top             =   7650
         Width           =   13995
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Index           =   2
            Left            =   540
            TabIndex        =   102
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
            MICON           =   "frmMstEmployee.frx":24A89
            PICN            =   "frmMstEmployee.frx":24AA5
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
            Index           =   2
            Left            =   1560
            TabIndex        =   103
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
            MICON           =   "frmMstEmployee.frx":25B37
            PICN            =   "frmMstEmployee.frx":25B53
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
            Index           =   2
            Left            =   2580
            TabIndex        =   104
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
            MICON           =   "frmMstEmployee.frx":26BE5
            PICN            =   "frmMstEmployee.frx":26C01
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
            Index           =   2
            Left            =   3600
            TabIndex        =   105
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
            MICON           =   "frmMstEmployee.frx":27C93
            PICN            =   "frmMstEmployee.frx":27CAF
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
            Index           =   2
            Left            =   4620
            TabIndex        =   106
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
            MICON           =   "frmMstEmployee.frx":28D41
            PICN            =   "frmMstEmployee.frx":28D5D
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
      Begin VB.Frame fra_entry_education 
         Height          =   3075
         Left            =   -74730
         TabIndex        =   94
         Top             =   4500
         Width           =   13965
         Begin VB.TextBox txt_edu_country 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5820
            TabIndex        =   53
            Top             =   2490
            Width           =   3705
         End
         Begin VB.TextBox txt_edu_majors 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            TabIndex        =   50
            Top             =   1440
            Width           =   3705
         End
         Begin VB.ComboBox cbo_edu_level 
            Height          =   315
            ItemData        =   "frmMstEmployee.frx":29DEF
            Left            =   5820
            List            =   "frmMstEmployee.frx":29E17
            TabIndex        =   48
            Text            =   "Pra TK"
            Top             =   1080
            Width           =   1545
         End
         Begin VB.TextBox txt_edu_school 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            TabIndex        =   51
            Top             =   1800
            Width           =   3705
         End
         Begin VB.TextBox txt_edu_city 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5820
            TabIndex        =   52
            Top             =   2160
            Width           =   3705
         End
         Begin MSComCtl2.DTPicker DTPicker_edu_start 
            Height          =   315
            Left            =   5820
            TabIndex        =   44
            Top             =   360
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_edu_end 
            Height          =   315
            Left            =   5820
            TabIndex        =   46
            Top             =   720
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NEGARA"
            Height          =   195
            Left            =   5010
            TabIndex        =   108
            Top             =   2490
            Width           =   675
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TAHUN LULUS*"
            Height          =   195
            Left            =   4455
            TabIndex        =   107
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JURUSAN"
            Height          =   195
            Left            =   4905
            TabIndex        =   99
            Top             =   1440
            Width           =   765
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "LEVEL*"
            Height          =   195
            Left            =   3840
            TabIndex        =   98
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TAHUN MASUK*"
            Height          =   195
            Left            =   4395
            TabIndex        =   97
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SEKOLAH/UNIVERSITAS*"
            Height          =   195
            Left            =   3705
            TabIndex        =   96
            Top             =   1800
            Width           =   1965
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KOTA"
            Height          =   195
            Left            =   5235
            TabIndex        =   95
            Top             =   2160
            Width           =   435
         End
      End
      Begin VB.Frame fra_entry_family 
         Height          =   3645
         Left            =   -74730
         TabIndex        =   84
         Top             =   3900
         Width           =   13965
         Begin VB.CheckBox chk_fams_address 
            Caption         =   "SAMA DENGAN KARYAWAN"
            Height          =   195
            Left            =   9600
            TabIndex        =   93
            Top             =   2700
            Width           =   3285
         End
         Begin VB.TextBox txt_fams_address 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   5820
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   2700
            Width           =   3705
         End
         Begin VB.TextBox txt_fams_employment 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5820
            TabIndex        =   42
            Top             =   2370
            Width           =   3705
         End
         Begin VB.TextBox txt_fams_edu 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            TabIndex        =   41
            Top             =   2010
            Width           =   3705
         End
         Begin VB.ComboBox cbo_fams_sex 
            Height          =   315
            ItemData        =   "frmMstEmployee.frx":29E53
            Left            =   5820
            List            =   "frmMstEmployee.frx":29E5D
            TabIndex        =   40
            Text            =   "Male"
            Top             =   1650
            Width           =   1545
         End
         Begin VB.ComboBox cbo_fams_rel 
            Height          =   315
            ItemData        =   "frmMstEmployee.frx":29E6F
            Left            =   5820
            List            =   "frmMstEmployee.frx":29E82
            TabIndex        =   38
            Text            =   "Suami"
            Top             =   930
            Width           =   1545
         End
         Begin VB.TextBox txt_family_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5820
            TabIndex        =   37
            Top             =   570
            Width           =   3705
         End
         Begin MSComCtl2.DTPicker DTPicker_fams_birth 
            Height          =   315
            Left            =   5820
            TabIndex        =   39
            Top             =   1290
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ALAMAT"
            Height          =   195
            Left            =   4935
            TabIndex        =   92
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PEKERJAAN"
            Height          =   195
            Left            =   4650
            TabIndex        =   91
            Top             =   2370
            Width           =   945
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PENDIDIKAN"
            Height          =   195
            Left            =   4605
            TabIndex        =   90
            Top             =   2010
            Width           =   990
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JENIS KELAMIN*"
            Height          =   195
            Left            =   4335
            TabIndex        =   89
            Top             =   1650
            Width           =   1260
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL. LAHIR"
            Height          =   195
            Left            =   4725
            TabIndex        =   88
            Top             =   1290
            Width           =   885
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "HUBUNGAN KELUARGA*"
            Height          =   195
            Left            =   3705
            TabIndex        =   86
            Top             =   930
            Width           =   1905
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA*"
            Height          =   195
            Left            =   5100
            TabIndex        =   85
            Top             =   570
            Width           =   525
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74730
         TabIndex        =   78
         Top             =   7650
         Width           =   13995
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Index           =   1
            Left            =   540
            TabIndex        =   79
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
            MICON           =   "frmMstEmployee.frx":29EAE
            PICN            =   "frmMstEmployee.frx":29ECA
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
            Index           =   1
            Left            =   1560
            TabIndex        =   80
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
            MICON           =   "frmMstEmployee.frx":2AF5C
            PICN            =   "frmMstEmployee.frx":2AF78
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
            Index           =   1
            Left            =   2580
            TabIndex        =   81
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
            MICON           =   "frmMstEmployee.frx":2C00A
            PICN            =   "frmMstEmployee.frx":2C026
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
            Index           =   1
            Left            =   3600
            TabIndex        =   82
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
            MICON           =   "frmMstEmployee.frx":2D0B8
            PICN            =   "frmMstEmployee.frx":2D0D4
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
            Index           =   1
            Left            =   4620
            TabIndex        =   83
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
            MICON           =   "frmMstEmployee.frx":2E166
            PICN            =   "frmMstEmployee.frx":2E182
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
      Begin VB.Frame fra_status_emp 
         Caption         =   "List Employee Status"
         Height          =   585
         Left            =   11520
         TabIndex        =   74
         Top             =   390
         Width           =   2925
         Begin VB.OptionButton optNotActive 
            Caption         =   "Not Active"
            Height          =   225
            Left            =   1500
            TabIndex        =   76
            Top             =   270
            Width           =   1245
         End
         Begin VB.OptionButton optActive 
            Caption         =   "Active"
            Height          =   225
            Left            =   240
            TabIndex        =   75
            Top             =   270
            Width           =   855
         End
      End
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   3180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   71
         Top             =   420
         Width           =   3855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   270
         TabIndex        =   49
         Top             =   7680
         Width           =   14175
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Index           =   0
            Left            =   540
            TabIndex        =   54
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
            MICON           =   "frmMstEmployee.frx":2F214
            PICN            =   "frmMstEmployee.frx":2F230
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
            Index           =   0
            Left            =   1560
            TabIndex        =   56
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
            MICON           =   "frmMstEmployee.frx":302C2
            PICN            =   "frmMstEmployee.frx":302DE
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
            Index           =   0
            Left            =   2580
            TabIndex        =   58
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
            MICON           =   "frmMstEmployee.frx":31370
            PICN            =   "frmMstEmployee.frx":3138C
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
            Index           =   0
            Left            =   3600
            TabIndex        =   69
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
            MICON           =   "frmMstEmployee.frx":3241E
            PICN            =   "frmMstEmployee.frx":3243A
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
            Index           =   0
            Left            =   4620
            TabIndex        =   70
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
            MICON           =   "frmMstEmployee.frx":334CC
            PICN            =   "frmMstEmployee.frx":334E8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdPrint 
            Height          =   705
            Left            =   8880
            TabIndex        =   269
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Cetak"
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
            MICON           =   "frmMstEmployee.frx":3457A
            PICN            =   "frmMstEmployee.frx":34596
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
            Left            =   9900
            TabIndex        =   270
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
            MICON           =   "frmMstEmployee.frx":35628
            PICN            =   "frmMstEmployee.frx":35644
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
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Left            =   1410
         OleObjectBlob   =   "frmMstEmployee.frx":366D6
         TabIndex        =   36
         Top             =   420
         Width           =   1695
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Family 
         Height          =   6915
         Left            =   -74730
         TabIndex        =   87
         Top             =   630
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   12197
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
         Columns(2).Caption=   "NAMA"
         Columns(2).DataField=   "name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "HUBUNGAN KELUARGA"
         Columns(3).DataField=   "nm_rel"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "JENIS KELAMIN"
         Columns(4).DataField=   "jenkel"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "PENDIDIKAN"
         Columns(5).DataField=   "education"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "PEKERJAAN"
         Columns(6).DataField=   "employment"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "ALAMAT"
         Columns(7).DataField=   "address"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=6562"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6482"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=4392"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=4313"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=3096"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=3016"
         Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=4339"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=4260"
         Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(33)=   "Column(6).Width=4577"
         Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=4498"
         Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(38)=   "Column(7).Width=6985"
         Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=6906"
         Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
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
         Caption         =   "DAFTAR KELUARGA"
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
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Education 
         Height          =   6915
         Left            =   -74730
         TabIndex        =   100
         Top             =   660
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   12197
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
         Columns(2).Caption=   "TAHUN MASUK"
         Columns(2).DataField=   "start_year"
         Columns(2).NumberFormat=   "yyyy-MM-dd"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "TAHUN LULUS"
         Columns(3).DataField=   "end_year"
         Columns(3).NumberFormat=   "yyyy-MM-dd"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "LEVEL"
         Columns(4).DataField=   "nm_jenjang"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "JURUSAN"
         Columns(5).DataField=   "jurusan"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "SEKOLAH/UNIVERSITAS"
         Columns(6).DataField=   "school"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "KOTA"
         Columns(7).DataField=   "city"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "NEGARA"
         Columns(8).DataField=   "country"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2461"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2381"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2408"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2328"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=2143"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2064"
         Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=3836"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=3757"
         Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(33)=   "Column(6).Width=5477"
         Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=5398"
         Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(38)=   "Column(7).Width=4048"
         Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=3969"
         Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(43)=   "Column(8).Width=3228"
         Splits(0)._ColumnProps(44)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(8)._WidthInPix=3149"
         Splits(0)._ColumnProps(46)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(47)=   "Column(8).Order=9"
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
         Caption         =   "DAFTAR PENDIDIKAN"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Skill 
         Height          =   6915
         Left            =   -74730
         TabIndex        =   117
         Top             =   660
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   12197
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
         Columns(2).Caption=   "KEAHLIAN"
         Columns(2).DataField=   "skill"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "LEVEL"
         Columns(3).DataField=   "nm_level"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=5477"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5398"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=15372"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=15293"
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
         Caption         =   "DAFTAR KEAHLIAN"
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
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=74,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=78,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Job 
         Height          =   6915
         Left            =   -74700
         TabIndex        =   133
         Top             =   630
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   12197
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
         Columns(2).Caption=   "TGL. MASUK"
         Columns(2).DataField=   "start_working"
         Columns(2).NumberFormat=   "yyyy-MM-dd"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "TGL. KELUAR"
         Columns(3).DataField=   "end_working"
         Columns(3).NumberFormat=   "yyyy-MM-dd"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "PERUSAHAAN"
         Columns(4).DataField=   "company_name"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "BIDANG USAHA"
         Columns(5).DataField=   "usaha"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DEPARTEMEN"
         Columns(6).DataField=   "department_name"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "JABATAN TERAKHIR"
         Columns(7).DataField=   "last_title"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "UPAH TERAKHIR"
         Columns(8).DataField=   "last_salary"
         Columns(8).NumberFormat=   "Standard"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "ALASAN KELUAR"
         Columns(9).DataField=   "reason_stop_working"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2461"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2381"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2408"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2328"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=4286"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=4207"
         Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=4313"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=4233"
         Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(33)=   "Column(6).Width=2963"
         Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=2884"
         Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(38)=   "Column(7).Width=2778"
         Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=2699"
         Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(43)=   "Column(8).Width=3201"
         Splits(0)._ColumnProps(44)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(8)._WidthInPix=3122"
         Splits(0)._ColumnProps(46)=   "Column(8)._ColStyle=514"
         Splits(0)._ColumnProps(47)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(48)=   "Column(9).Width=5530"
         Splits(0)._ColumnProps(49)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(50)=   "Column(9)._WidthInPix=5450"
         Splits(0)._ColumnProps(51)=   "Column(9)._ColStyle=516"
         Splits(0)._ColumnProps(52)=   "Column(9).Order=10"
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
         Caption         =   "DAFTAR PENGALAMAN KERJA"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=3"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=3"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3045
         Left            =   -74670
         TabIndex        =   137
         Top             =   630
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   5371
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
         Columns(2).Caption=   "DATE"
         Columns(2).DataField=   "date_grade"
         Columns(2).NumberFormat=   "yyyy-MM-dd"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "GRADE"
         Columns(3).DataField=   "grade_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DESCRIPTION"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3863"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3784"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=6562"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=6482"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=13150"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=13070"
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
         Caption         =   "LIST OF GRADE"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=3"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   3645
         Left            =   -74670
         TabIndex        =   138
         Top             =   3840
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   6429
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
         Columns(2).Caption=   "DATE"
         Columns(2).DataField=   "date_grade"
         Columns(2).NumberFormat=   "yyyy-MM-dd"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "GRADE"
         Columns(3).DataField=   "grade_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DESCRIPTION"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3863"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3784"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=6562"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=6482"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=13150"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=13070"
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
         Caption         =   "LIST OF GRADE"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=3"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Contract 
         Height          =   6915
         Left            =   -74730
         TabIndex        =   210
         Top             =   630
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   12197
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
         Columns(2).Caption=   "TGL. AWAL"
         Columns(2).DataField=   "start_date"
         Columns(2).NumberFormat=   "yyyy-MM-dd"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "TGL. AKHIR"
         Columns(3).DataField=   "end_date"
         Columns(3).NumberFormat=   "yyyy-MM-dd"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "NO. KONTRAK"
         Columns(4).DataField=   "no_contract"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "KETERANGAN"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2461"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2381"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2408"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2328"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=7938"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=7858"
         Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=10769"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=10689"
         Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
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
         Caption         =   "DAFTAR CONTRACT"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=3"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=3"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Training 
         Height          =   6825
         Left            =   -74730
         TabIndex        =   230
         Top             =   690
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   12039
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
         Columns(2).Caption=   "TGL. MULAI"
         Columns(2).DataField=   "start_date"
         Columns(2).NumberFormat=   "yyyy-MM-dd"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "TGL. BERAKHIR"
         Columns(3).DataField=   "end_date"
         Columns(3).NumberFormat=   "yyyy-MM-dd"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "NAMA TRAINING"
         Columns(4).DataField=   "material"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "PENYELENGGARA"
         Columns(5).DataField=   "organizer"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "TEMPAT"
         Columns(6).DataField=   "place"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "NILAI"
         Columns(7).DataField=   "value"
         Columns(7).NumberFormat=   "General Number"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2461"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2381"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2408"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2328"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=5450"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=5371"
         Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=5265"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=5186"
         Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(33)=   "Column(6).Width=5450"
         Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=5371"
         Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(38)=   "Column(7).Width=2593"
         Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=2514"
         Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=513"
         Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
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
         Caption         =   "DAFTAR TRAINING"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=3"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=3"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Grade 
         Height          =   6915
         Left            =   -74700
         TabIndex        =   245
         Top             =   600
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   12197
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
         Columns(2).DataField=   "date_grade"
         Columns(2).NumberFormat=   "yyyy-MM-dd"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "GRADE"
         Columns(3).DataField=   "grade_name"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3863"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3784"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=6562"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=6482"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=13150"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=13070"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=3"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Title 
         Height          =   6915
         Left            =   -74760
         TabIndex        =   260
         Top             =   660
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   12197
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
         Columns(2).DataField=   "date_title"
         Columns(2).NumberFormat=   "yyyy-MM-dd"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "JABATAN"
         Columns(3).DataField=   "title_name"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3863"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3784"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=6562"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=6482"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=13150"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=13070"
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
         Caption         =   "DAFTAR JABATAN"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=3"
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
      Begin VB.Frame fra_entry_employee 
         Height          =   6705
         Left            =   270
         TabIndex        =   140
         Top             =   960
         Width           =   14175
         Begin VB.TextBox txt_pict_location 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   13230
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   157
            Top             =   1260
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.TextBox txtNoKTP 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   11070
            TabIndex        =   30
            Top             =   3180
            Width           =   2955
         End
         Begin VB.TextBox txt_address 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   6120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   3930
            Width           =   2775
         End
         Begin VB.TextBox txt_employee_code 
            Height          =   315
            Left            =   3930
            TabIndex        =   156
            Top             =   210
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txt_jamsostek 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            TabIndex        =   19
            Top             =   4410
            Width           =   2775
         End
         Begin VB.TextBox txt_email 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   11070
            TabIndex        =   28
            Top             =   2460
            Width           =   2955
         End
         Begin VB.TextBox txt_number_of_children 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            TabIndex        =   13
            Top             =   2010
            Width           =   1455
         End
         Begin VB.TextBox txt_npwp 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            TabIndex        =   14
            Top             =   2370
            Width           =   2775
         End
         Begin VB.TextBox txt_employee_nick_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   930
            Width           =   2655
         End
         Begin VB.ComboBox cbo_religion 
            Height          =   315
            ItemData        =   "frmMstEmployee.frx":3863C
            Left            =   6120
            List            =   "frmMstEmployee.frx":38652
            TabIndex        =   11
            Text            =   "Islam"
            Top             =   1290
            Width           =   2805
         End
         Begin VB.TextBox txt_description 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   11070
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   4260
            Width           =   3015
         End
         Begin VB.TextBox txt_age 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   8400
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   155
            Top             =   570
            Width           =   495
         End
         Begin VB.Frame Frame3 
            Caption         =   "Status Kepegawaian"
            Enabled         =   0   'False
            Height          =   585
            Left            =   9240
            TabIndex        =   154
            Top             =   4740
            Width           =   4845
            Begin VB.OptionButton opt_not_active 
               Caption         =   "TIDAK AKTIF"
               Height          =   255
               Left            =   1320
               TabIndex        =   33
               Top             =   300
               Width           =   1515
            End
            Begin VB.OptionButton opt_active 
               Caption         =   "AKTIF"
               Height          =   255
               Left            =   240
               TabIndex        =   32
               Top             =   300
               Value           =   -1  'True
               Width           =   1035
            End
         End
         Begin VB.TextBox txt_working_time 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   12090
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   153
            Top             =   1380
            Width           =   495
         End
         Begin VB.TextBox txt_bank_account 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            TabIndex        =   22
            Top             =   5910
            Width           =   2775
         End
         Begin VB.TextBox txt_phone_number 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   11070
            TabIndex        =   27
            Top             =   2100
            Width           =   2955
         End
         Begin VB.ComboBox cbo_marital_status 
            Height          =   315
            ItemData        =   "frmMstEmployee.frx":38696
            Left            =   6120
            List            =   "frmMstEmployee.frx":386A0
            TabIndex        =   12
            Text            =   "Single"
            Top             =   1650
            Width           =   2805
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
            Height          =   375
            Left            =   1560
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   152
            Top             =   1650
            Width           =   2655
         End
         Begin VB.TextBox txt_nik 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   0
            Top             =   210
            Width           =   2655
         End
         Begin VB.TextBox txt_place_of_birth 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            TabIndex        =   8
            Top             =   210
            Width           =   2775
         End
         Begin VB.TextBox txtAlamat 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   6120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   3450
            Width           =   2775
         End
         Begin VB.TextBox txt_employee_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Top             =   570
            Width           =   2655
         End
         Begin VB.ComboBox cbo_sex 
            Height          =   315
            ItemData        =   "frmMstEmployee.frx":386B5
            Left            =   6120
            List            =   "frmMstEmployee.frx":386BF
            TabIndex        =   10
            Text            =   "Male"
            Top             =   930
            Width           =   2805
         End
         Begin VB.TextBox txt_title_name 
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
            Left            =   1560
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   151
            Top             =   2430
            Width           =   2655
         End
         Begin VB.TextBox txt_level_name 
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
            Left            =   1560
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   150
            Top             =   3990
            Width           =   2655
         End
         Begin VB.TextBox txt_grade_name 
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
            Left            =   1560
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   149
            Top             =   3210
            Width           =   2655
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
            Left            =   1560
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   148
            Top             =   4770
            Width           =   2655
         End
         Begin VB.ComboBox cbo_tax_method 
            Height          =   315
            ItemData        =   "frmMstEmployee.frx":386D1
            Left            =   6120
            List            =   "frmMstEmployee.frx":386DE
            TabIndex        =   15
            Text            =   "Gross (Pajak Ditanggung Karyawan)"
            Top             =   2730
            Width           =   2805
         End
         Begin VB.TextBox txt_bank_name 
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
            Left            =   6120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   147
            Top             =   5490
            Width           =   2775
         End
         Begin VB.TextBox txt_account_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            TabIndex        =   23
            Top             =   6270
            Width           =   2775
         End
         Begin VB.ComboBox cbo_nationality 
            Height          =   315
            ItemData        =   "frmMstEmployee.frx":3875C
            Left            =   11070
            List            =   "frmMstEmployee.frx":38766
            TabIndex        =   29
            Text            =   "WNI"
            Top             =   2820
            Width           =   2985
         End
         Begin VB.TextBox txt_last_edu 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   11070
            TabIndex        =   146
            Top             =   3540
            Width           =   2955
         End
         Begin VB.TextBox txt_last_emp 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   11070
            TabIndex        =   145
            Top             =   3900
            Width           =   2955
         End
         Begin VB.PictureBox pic 
            Height          =   1845
            Left            =   12630
            ScaleHeight     =   1785
            ScaleWidth      =   1335
            TabIndex        =   144
            Top             =   210
            Width           =   1395
            Begin VB.Image img 
               Height          =   1485
               Left            =   120
               Stretch         =   -1  'True
               Top             =   150
               Width           =   1095
            End
         End
         Begin prj_panji.vbButton cmd_brows_pict 
            Height          =   585
            Left            =   11070
            TabIndex        =   24
            Top             =   360
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
            MICON           =   "frmMstEmployee.frx":38774
            PICN            =   "frmMstEmployee.frx":38790
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DTPicker_birth 
            Height          =   315
            Left            =   6120
            TabIndex        =   9
            Top             =   570
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_start_working 
            Height          =   315
            Left            =   11070
            TabIndex        =   25
            Top             =   1020
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_appointment 
            Height          =   315
            Left            =   11070
            TabIndex        =   26
            Top             =   1740
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_division 
            Height          =   375
            Left            =   1560
            OleObjectBlob   =   "frmMstEmployee.frx":39822
            TabIndex        =   3
            Top             =   1290
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_tglnpwp 
            Height          =   315
            Left            =   6120
            TabIndex        =   16
            Top             =   3090
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_jstk 
            Height          =   315
            Left            =   6120
            TabIndex        =   20
            Top             =   4770
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39270
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_title 
            Height          =   375
            Left            =   1560
            OleObjectBlob   =   "frmMstEmployee.frx":3B789
            TabIndex        =   4
            Top             =   2070
            Width           =   1695
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_level 
            Height          =   375
            Left            =   1560
            OleObjectBlob   =   "frmMstEmployee.frx":3D6E5
            TabIndex        =   6
            Top             =   3630
            Width           =   1695
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_grade 
            Height          =   375
            Left            =   1560
            OleObjectBlob   =   "frmMstEmployee.frx":3F641
            TabIndex        =   5
            Top             =   2850
            Width           =   1695
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_status 
            Height          =   375
            Left            =   1560
            OleObjectBlob   =   "frmMstEmployee.frx":4159D
            TabIndex        =   7
            Top             =   4410
            Width           =   1695
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_bank 
            Height          =   375
            Left            =   6120
            OleObjectBlob   =   "frmMstEmployee.frx":434FA
            TabIndex        =   21
            Top             =   5130
            Width           =   1515
         End
         Begin VB.Frame fra_not_active 
            Height          =   1305
            Left            =   9240
            TabIndex        =   141
            Top             =   5250
            Visible         =   0   'False
            Width           =   4845
            Begin VB.TextBox txt_end_working_reason 
               Appearance      =   0  'Flat
               Height          =   465
               Left            =   1560
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   35
               Top             =   720
               Width           =   2655
            End
            Begin MSComCtl2.DTPicker DTPicker_end_working 
               Height          =   315
               Left            =   1560
               TabIndex        =   34
               Top             =   330
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               MousePointer    =   99
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   96141315
               CurrentDate     =   39270
            End
            Begin VB.Label lbl_reason_description 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "ALASAN KELUAR"
               Height          =   195
               Left            =   150
               TabIndex        =   143
               Top             =   750
               Width           =   1320
            End
            Begin VB.Label lbl_reason_end_working 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "TGL. KELUAR"
               Height          =   195
               Left            =   420
               TabIndex        =   142
               Top             =   360
               Width           =   1050
            End
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            Caption         =   "(Max 100 Kb)"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   9630
            TabIndex        =   194
            Top             =   450
            Width           =   1245
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NO. IDENTITAS"
            Height          =   195
            Left            =   9720
            TabIndex        =   193
            Top             =   3180
            Width           =   1185
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ALAMAT KARY."
            Height          =   195
            Left            =   4560
            TabIndex        =   192
            Top             =   3480
            Width           =   1470
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL REG. NPWP"
            Height          =   195
            Left            =   4740
            TabIndex        =   191
            Top             =   3120
            Width           =   1290
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NO. JAMSOSTEK"
            Height          =   195
            Left            =   4560
            TabIndex        =   190
            Top             =   4440
            Width           =   1470
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EMAIL"
            Height          =   195
            Left            =   9270
            TabIndex        =   189
            Top             =   2460
            Width           =   1620
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JUMLAH ANAK*"
            Height          =   195
            Left            =   4845
            TabIndex        =   188
            Top             =   2010
            Width           =   1185
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "N P W P"
            Height          =   195
            Left            =   4560
            TabIndex        =   187
            Top             =   2370
            Width           =   1470
         End
         Begin VB.Label lbl_employee_nick_name 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA PENDEK*"
            Height          =   195
            Left            =   210
            TabIndex        =   186
            Top             =   930
            Width           =   1260
         End
         Begin VB.Label lbl_religion 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "AGAMA*"
            Height          =   195
            Left            =   5385
            TabIndex        =   185
            Top             =   1290
            Width           =   630
         End
         Begin VB.Label lbl_description 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   9795
            TabIndex        =   184
            Top             =   4260
            Width           =   1110
         End
         Begin VB.Label lbl_appointment 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PENGANGKATAN"
            Height          =   195
            Left            =   9555
            TabIndex        =   183
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label lbl_age 
            AutoSize        =   -1  'True
            Caption         =   "AGE (Y)*"
            ForeColor       =   &H80000002&
            Height          =   195
            Left            =   7680
            TabIndex        =   182
            Top             =   570
            Width           =   630
         End
         Begin VB.Label lbl_working_age 
            AutoSize        =   -1  'True
            Caption         =   "W. AGE (Y)*"
            ForeColor       =   &H80000002&
            Height          =   195
            Left            =   11070
            TabIndex        =   181
            Top             =   1410
            Width           =   915
         End
         Begin VB.Label lbl_start_working 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL. MASUK"
            Height          =   195
            Left            =   9915
            TabIndex        =   180
            Top             =   1050
            Width           =   975
         End
         Begin VB.Label lbl_bank_account 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NO. REK."
            Height          =   195
            Left            =   4560
            TabIndex        =   179
            Top             =   5910
            Width           =   1485
         End
         Begin VB.Label lbl_phone_number 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NO. TELP."
            Height          =   195
            Left            =   10110
            TabIndex        =   178
            Top             =   2100
            Width           =   780
         End
         Begin VB.Label lbl_marital_status 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "STATUS KAWIN*"
            Height          =   195
            Left            =   4725
            TabIndex        =   177
            Top             =   1650
            Width           =   1290
         End
         Begin VB.Label lbl_sex 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JENIS KELAMIN*"
            Height          =   195
            Left            =   4560
            TabIndex        =   176
            Top             =   930
            Width           =   1470
         End
         Begin VB.Label lbl_division 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DIVISI*"
            Height          =   195
            Left            =   930
            TabIndex        =   175
            Top             =   1290
            Width           =   525
         End
         Begin VB.Label lbl_employee_name 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA LENGKAP*"
            Height          =   195
            Left            =   60
            TabIndex        =   174
            Top             =   570
            Width           =   1410
         End
         Begin VB.Label lbl_address 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ALAMAT NPWP"
            Height          =   195
            Left            =   4560
            TabIndex        =   173
            Top             =   3930
            Width           =   1455
         End
         Begin VB.Label lbl_employee_code 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE KARY.*"
            Height          =   195
            Left            =   360
            TabIndex        =   172
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label lbl_date_of_birth 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TANGGAL LAHIR*"
            Height          =   195
            Left            =   4680
            TabIndex        =   171
            Top             =   570
            Width           =   1350
         End
         Begin VB.Label lbl_place_of_birth 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TEMPAT LAHIR*"
            Height          =   195
            Left            =   4800
            TabIndex        =   170
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label lbl_title 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JABATAN*"
            Height          =   195
            Left            =   675
            TabIndex        =   169
            Top             =   2100
            Width           =   780
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "STATUS*"
            Height          =   195
            Left            =   330
            TabIndex        =   168
            Top             =   4440
            Width           =   1125
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "LEVEL*"
            Height          =   195
            Left            =   360
            TabIndex        =   167
            Top             =   3660
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "GRADE*"
            Height          =   195
            Left            =   360
            TabIndex        =   166
            Top             =   2880
            Width           =   1110
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "METODE PAJAK"
            Height          =   195
            Left            =   4815
            TabIndex        =   165
            Top             =   2730
            Width           =   1230
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL REG. JSTK"
            Height          =   195
            Left            =   4560
            TabIndex        =   164
            Top             =   4800
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA BANK"
            Height          =   195
            Left            =   4560
            TabIndex        =   163
            Top             =   5160
            Width           =   1455
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA REKENING"
            Height          =   195
            Left            =   4560
            TabIndex        =   162
            Top             =   6270
            Width           =   1470
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KEWARGANEGARAAN"
            Height          =   195
            Left            =   9180
            TabIndex        =   161
            Top             =   2820
            Width           =   1725
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PEND. TERAKHIR"
            Height          =   195
            Left            =   9540
            TabIndex        =   160
            Top             =   3540
            Width           =   1365
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PEKERJAAN TERAKHIR"
            Height          =   195
            Left            =   9090
            TabIndex        =   159
            Top             =   3900
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "GAMBAR"
            Height          =   195
            Left            =   9300
            TabIndex        =   158
            Top             =   240
            Width           =   1590
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Employee 
         Height          =   6705
         Left            =   270
         TabIndex        =   73
         Top             =   960
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   11827
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
         Columns(1).Caption=   "PERUSAHAAN"
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
         Columns(6).Caption=   "KODE KARY."
         Columns(6).DataField=   "nik"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "EMPLOYEE CODE"
         Columns(7).DataField=   "employee_code"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "NAMA KARY."
         Columns(8).DataField=   "employee_name"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "NAMA PENDEK"
         Columns(9).DataField=   "employee_nick_name"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   4
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "AKTIF"
         Columns(10).DataField=   "flag_active"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "EMAIL"
         Columns(11).DataField=   "email"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "TGL. LAHIR"
         Columns(12).DataField=   "date_birth"
         Columns(12).NumberFormat=   "FormatText Event"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "TEMPAT LAHIR"
         Columns(13).DataField=   "place_birth"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   16
         Columns(14)._MaxComboItems=   5
         Columns(14).ValueItems(0)._DefaultItem=   0
         Columns(14).ValueItems(0).Value=   "0"
         Columns(14).ValueItems(0).Value.vt=   8
         Columns(14).ValueItems(0).DisplayValue=   "Female"
         Columns(14).ValueItems(0).DisplayValue.vt=   8
         Columns(14).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(14).ValueItems(1)._DefaultItem=   0
         Columns(14).ValueItems(1).Value=   "1"
         Columns(14).ValueItems(1).Value.vt=   8
         Columns(14).ValueItems(1).DisplayValue=   "Male"
         Columns(14).ValueItems(1).DisplayValue.vt=   8
         Columns(14).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(14).ValueItems.Count=   2
         Columns(14).Caption=   "JENIS KELAMIN"
         Columns(14).DataField=   "sex"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   16
         Columns(15)._MaxComboItems=   5
         Columns(15).ValueItems(0)._DefaultItem=   0
         Columns(15).ValueItems(0).Value=   "0"
         Columns(15).ValueItems(0).Value.vt=   8
         Columns(15).ValueItems(0).DisplayValue=   "Single"
         Columns(15).ValueItems(0).DisplayValue.vt=   8
         Columns(15).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(15).ValueItems(1)._DefaultItem=   0
         Columns(15).ValueItems(1).Value=   "1"
         Columns(15).ValueItems(1).Value.vt=   8
         Columns(15).ValueItems(1).DisplayValue=   "Married"
         Columns(15).ValueItems(1).DisplayValue.vt=   8
         Columns(15).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(15).ValueItems(2)._DefaultItem=   0
         Columns(15).ValueItems(2).Value=   "2"
         Columns(15).ValueItems(2).Value.vt=   8
         Columns(15).ValueItems(2).DisplayValue=   "Widow"
         Columns(15).ValueItems(2).DisplayValue.vt=   8
         Columns(15).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(15).ValueItems(3)._DefaultItem=   0
         Columns(15).ValueItems(3).Value=   "3"
         Columns(15).ValueItems(3).Value.vt=   8
         Columns(15).ValueItems(3).DisplayValue=   "Widower"
         Columns(15).ValueItems(3).DisplayValue.vt=   8
         Columns(15).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
         Columns(15).ValueItems.Count=   4
         Columns(15).Caption=   "STATUS"
         Columns(15).DataField=   "marital_status"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "ALAMAT"
         Columns(16).DataField=   "emp_address"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "NO. TELP."
         Columns(17).DataField=   "phone_number"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "NO. REK."
         Columns(18).DataField=   "bank_account"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "TGL. MASUK"
         Columns(19).DataField=   "start_working"
         Columns(19).NumberFormat=   "FormatText Event"
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).Caption=   "TITLE CODE"
         Columns(20).DataField=   "title_code"
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(21)._VlistStyle=   0
         Columns(21)._MaxComboItems=   5
         Columns(21).Caption=   "JABATAN"
         Columns(21).DataField=   "title_name"
         Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(22)._VlistStyle=   0
         Columns(22)._MaxComboItems=   5
         Columns(22).Caption=   "TGL. KELUAR"
         Columns(22).DataField=   "end_working"
         Columns(22).NumberFormat=   "FormatText Event"
         Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(23)._VlistStyle=   0
         Columns(23)._MaxComboItems=   5
         Columns(23).Caption=   "KETERANGAN"
         Columns(23).DataField=   "reason"
         Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(24)._VlistStyle=   0
         Columns(24)._MaxComboItems=   5
         Columns(24).Caption=   "Picture"
         Columns(24).DataField=   "picture"
         Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   25
         Splits(0)._UserFlags=   0
         Splits(0).SizeMode=   1
         Splits(0).Size  =   4004.788
         Splits(0).Size.vt=   4
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   3
         Splits(0).DividerColor=   13160660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=25"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1958"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1879"
         Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=3916"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3836"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=2064"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=1984"
         Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(23)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(25)=   "Column(3).Width=3545"
         Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=3466"
         Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=8708"
         Splits(0)._ColumnProps(29)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(31)=   "Column(4).Width=2037"
         Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=1958"
         Splits(0)._ColumnProps(34)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(36)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(38)=   "Column(5).Width=4022"
         Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=3942"
         Splits(0)._ColumnProps(41)=   "Column(5)._ColStyle=8708"
         Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(43)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(47)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(48)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(49)=   "Column(7).Width=1588"
         Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=1508"
         Splits(0)._ColumnProps(52)=   "Column(7).AllowSizing=0"
         Splits(0)._ColumnProps(53)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(54)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(55)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(56)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(57)=   "Column(8).Width=1588"
         Splits(0)._ColumnProps(58)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(8)._WidthInPix=1508"
         Splits(0)._ColumnProps(60)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(61)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(62)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(63)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(64)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(65)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(66)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(67)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(68)=   "Column(9).AllowSizing=0"
         Splits(0)._ColumnProps(69)=   "Column(9)._ColStyle=516"
         Splits(0)._ColumnProps(70)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(71)=   "Column(9).AllowFocus=0"
         Splits(0)._ColumnProps(72)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(73)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(74)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(76)=   "Column(10).AllowSizing=0"
         Splits(0)._ColumnProps(77)=   "Column(10)._ColStyle=513"
         Splits(0)._ColumnProps(78)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(79)=   "Column(10).AllowFocus=0"
         Splits(0)._ColumnProps(80)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(81)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(82)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(83)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(84)=   "Column(11).AllowSizing=0"
         Splits(0)._ColumnProps(85)=   "Column(11)._ColStyle=516"
         Splits(0)._ColumnProps(86)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(87)=   "Column(11).AllowFocus=0"
         Splits(0)._ColumnProps(88)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(89)=   "Column(12).Width=2064"
         Splits(0)._ColumnProps(90)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(91)=   "Column(12)._WidthInPix=1984"
         Splits(0)._ColumnProps(92)=   "Column(12).AllowSizing=0"
         Splits(0)._ColumnProps(93)=   "Column(12)._ColStyle=516"
         Splits(0)._ColumnProps(94)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(95)=   "Column(12).AllowFocus=0"
         Splits(0)._ColumnProps(96)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(97)=   "Column(13).Width=3016"
         Splits(0)._ColumnProps(98)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(99)=   "Column(13)._WidthInPix=2937"
         Splits(0)._ColumnProps(100)=   "Column(13).AllowSizing=0"
         Splits(0)._ColumnProps(101)=   "Column(13)._ColStyle=516"
         Splits(0)._ColumnProps(102)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(103)=   "Column(13).AllowFocus=0"
         Splits(0)._ColumnProps(104)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(105)=   "Column(14).Width=2037"
         Splits(0)._ColumnProps(106)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(107)=   "Column(14)._WidthInPix=1958"
         Splits(0)._ColumnProps(108)=   "Column(14).AllowSizing=0"
         Splits(0)._ColumnProps(109)=   "Column(14)._ColStyle=516"
         Splits(0)._ColumnProps(110)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(111)=   "Column(14).AllowFocus=0"
         Splits(0)._ColumnProps(112)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(113)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(114)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(115)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(116)=   "Column(15).AllowSizing=0"
         Splits(0)._ColumnProps(117)=   "Column(15)._ColStyle=516"
         Splits(0)._ColumnProps(118)=   "Column(15).Visible=0"
         Splits(0)._ColumnProps(119)=   "Column(15).AllowFocus=0"
         Splits(0)._ColumnProps(120)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(121)=   "Column(16).Width=2725"
         Splits(0)._ColumnProps(122)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(123)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(124)=   "Column(16).AllowSizing=0"
         Splits(0)._ColumnProps(125)=   "Column(16)._ColStyle=516"
         Splits(0)._ColumnProps(126)=   "Column(16).Visible=0"
         Splits(0)._ColumnProps(127)=   "Column(16).AllowFocus=0"
         Splits(0)._ColumnProps(128)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(129)=   "Column(16)._MinWidth=10"
         Splits(0)._ColumnProps(130)=   "Column(17).Width=2725"
         Splits(0)._ColumnProps(131)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(132)=   "Column(17)._WidthInPix=2646"
         Splits(0)._ColumnProps(133)=   "Column(17).AllowSizing=0"
         Splits(0)._ColumnProps(134)=   "Column(17)._ColStyle=516"
         Splits(0)._ColumnProps(135)=   "Column(17).Visible=0"
         Splits(0)._ColumnProps(136)=   "Column(17).AllowFocus=0"
         Splits(0)._ColumnProps(137)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(138)=   "Column(17)._MinWidth=54215968"
         Splits(0)._ColumnProps(139)=   "Column(18).Width=2725"
         Splits(0)._ColumnProps(140)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(141)=   "Column(18)._WidthInPix=2646"
         Splits(0)._ColumnProps(142)=   "Column(18).AllowSizing=0"
         Splits(0)._ColumnProps(143)=   "Column(18)._ColStyle=516"
         Splits(0)._ColumnProps(144)=   "Column(18).Visible=0"
         Splits(0)._ColumnProps(145)=   "Column(18).AllowFocus=0"
         Splits(0)._ColumnProps(146)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(147)=   "Column(19).Width=2725"
         Splits(0)._ColumnProps(148)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(149)=   "Column(19)._WidthInPix=2646"
         Splits(0)._ColumnProps(150)=   "Column(19).AllowSizing=0"
         Splits(0)._ColumnProps(151)=   "Column(19)._ColStyle=516"
         Splits(0)._ColumnProps(152)=   "Column(19).Visible=0"
         Splits(0)._ColumnProps(153)=   "Column(19).AllowFocus=0"
         Splits(0)._ColumnProps(154)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(155)=   "Column(19)._MinWidth=60129312"
         Splits(0)._ColumnProps(156)=   "Column(20).Width=2725"
         Splits(0)._ColumnProps(157)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(158)=   "Column(20)._WidthInPix=2646"
         Splits(0)._ColumnProps(159)=   "Column(20).AllowSizing=0"
         Splits(0)._ColumnProps(160)=   "Column(20)._ColStyle=516"
         Splits(0)._ColumnProps(161)=   "Column(20).Visible=0"
         Splits(0)._ColumnProps(162)=   "Column(20).AllowFocus=0"
         Splits(0)._ColumnProps(163)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(164)=   "Column(21).Width=2725"
         Splits(0)._ColumnProps(165)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(166)=   "Column(21)._WidthInPix=2646"
         Splits(0)._ColumnProps(167)=   "Column(21).AllowSizing=0"
         Splits(0)._ColumnProps(168)=   "Column(21)._ColStyle=516"
         Splits(0)._ColumnProps(169)=   "Column(21).Visible=0"
         Splits(0)._ColumnProps(170)=   "Column(21).AllowFocus=0"
         Splits(0)._ColumnProps(171)=   "Column(21).Order=22"
         Splits(0)._ColumnProps(172)=   "Column(21)._MinWidth=79702332"
         Splits(0)._ColumnProps(173)=   "Column(22).Width=2725"
         Splits(0)._ColumnProps(174)=   "Column(22).DividerColor=0"
         Splits(0)._ColumnProps(175)=   "Column(22)._WidthInPix=2646"
         Splits(0)._ColumnProps(176)=   "Column(22).AllowSizing=0"
         Splits(0)._ColumnProps(177)=   "Column(22)._ColStyle=516"
         Splits(0)._ColumnProps(178)=   "Column(22).Visible=0"
         Splits(0)._ColumnProps(179)=   "Column(22).AllowFocus=0"
         Splits(0)._ColumnProps(180)=   "Column(22).Order=23"
         Splits(0)._ColumnProps(181)=   "Column(22)._MinWidth=79914544"
         Splits(0)._ColumnProps(182)=   "Column(23).Width=2725"
         Splits(0)._ColumnProps(183)=   "Column(23).DividerColor=0"
         Splits(0)._ColumnProps(184)=   "Column(23)._WidthInPix=2646"
         Splits(0)._ColumnProps(185)=   "Column(23).AllowSizing=0"
         Splits(0)._ColumnProps(186)=   "Column(23)._ColStyle=516"
         Splits(0)._ColumnProps(187)=   "Column(23).Visible=0"
         Splits(0)._ColumnProps(188)=   "Column(23).AllowFocus=0"
         Splits(0)._ColumnProps(189)=   "Column(23).Order=24"
         Splits(0)._ColumnProps(190)=   "Column(23)._MinWidth=79789632"
         Splits(0)._ColumnProps(191)=   "Column(24).Width=2725"
         Splits(0)._ColumnProps(192)=   "Column(24).DividerColor=0"
         Splits(0)._ColumnProps(193)=   "Column(24)._WidthInPix=2646"
         Splits(0)._ColumnProps(194)=   "Column(24)._ColStyle=516"
         Splits(0)._ColumnProps(195)=   "Column(24).Visible=0"
         Splits(0)._ColumnProps(196)=   "Column(24).Order=25"
         Splits(1)._UserFlags=   0
         Splits(1).Size  =   2
         Splits(1).Size.vt=   2
         Splits(1).RecordSelectors=   0   'False
         Splits(1).RecordSelectorWidth=   503
         Splits(1)._SavedRecordSelectors=   0   'False
         Splits(1).ScrollBars=   3
         Splits(1).DividerColor=   13160660
         Splits(1).FilterBar=   -1  'True
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=25"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=1826"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1746"
         Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(9)=   "Column(1).Width=1826"
         Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=1746"
         Splits(1)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=516"
         Splits(1)._ColumnProps(14)=   "Column(1).Visible=0"
         Splits(1)._ColumnProps(15)=   "Column(1).AllowFocus=0"
         Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(17)=   "Column(2).Width=1720"
         Splits(1)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(19)=   "Column(2)._WidthInPix=1640"
         Splits(1)._ColumnProps(20)=   "Column(2).AllowSizing=0"
         Splits(1)._ColumnProps(21)=   "Column(2)._ColStyle=516"
         Splits(1)._ColumnProps(22)=   "Column(2).Visible=0"
         Splits(1)._ColumnProps(23)=   "Column(2).AllowFocus=0"
         Splits(1)._ColumnProps(24)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(25)=   "Column(3).Width=1720"
         Splits(1)._ColumnProps(26)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(27)=   "Column(3)._WidthInPix=1640"
         Splits(1)._ColumnProps(28)=   "Column(3).AllowSizing=0"
         Splits(1)._ColumnProps(29)=   "Column(3)._ColStyle=516"
         Splits(1)._ColumnProps(30)=   "Column(3).Visible=0"
         Splits(1)._ColumnProps(31)=   "Column(3).AllowFocus=0"
         Splits(1)._ColumnProps(32)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(33)=   "Column(4).Width=1508"
         Splits(1)._ColumnProps(34)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(35)=   "Column(4)._WidthInPix=1429"
         Splits(1)._ColumnProps(36)=   "Column(4).AllowSizing=0"
         Splits(1)._ColumnProps(37)=   "Column(4)._ColStyle=516"
         Splits(1)._ColumnProps(38)=   "Column(4).Visible=0"
         Splits(1)._ColumnProps(39)=   "Column(4).AllowFocus=0"
         Splits(1)._ColumnProps(40)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(41)=   "Column(4)._MinWidth=80002672"
         Splits(1)._ColumnProps(42)=   "Column(5).Width=1508"
         Splits(1)._ColumnProps(43)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(44)=   "Column(5)._WidthInPix=1429"
         Splits(1)._ColumnProps(45)=   "Column(5).AllowSizing=0"
         Splits(1)._ColumnProps(46)=   "Column(5)._ColStyle=516"
         Splits(1)._ColumnProps(47)=   "Column(5).Visible=0"
         Splits(1)._ColumnProps(48)=   "Column(5).AllowFocus=0"
         Splits(1)._ColumnProps(49)=   "Column(5).Order=6"
         Splits(1)._ColumnProps(50)=   "Column(5)._MinWidth=80001968"
         Splits(1)._ColumnProps(51)=   "Column(6).Width=2725"
         Splits(1)._ColumnProps(52)=   "Column(6).DividerColor=0"
         Splits(1)._ColumnProps(53)=   "Column(6)._WidthInPix=2646"
         Splits(1)._ColumnProps(54)=   "Column(6)._ColStyle=516"
         Splits(1)._ColumnProps(55)=   "Column(6).Order=7"
         Splits(1)._ColumnProps(56)=   "Column(6)._MinWidth=80000960"
         Splits(1)._ColumnProps(57)=   "Column(7).Width=1879"
         Splits(1)._ColumnProps(58)=   "Column(7).DividerColor=0"
         Splits(1)._ColumnProps(59)=   "Column(7)._WidthInPix=1799"
         Splits(1)._ColumnProps(60)=   "Column(7)._ColStyle=8708"
         Splits(1)._ColumnProps(61)=   "Column(7).Visible=0"
         Splits(1)._ColumnProps(62)=   "Column(7).Order=8"
         Splits(1)._ColumnProps(63)=   "Column(7)._MinWidth=80000960"
         Splits(1)._ColumnProps(64)=   "Column(8).Width=3678"
         Splits(1)._ColumnProps(65)=   "Column(8).DividerColor=0"
         Splits(1)._ColumnProps(66)=   "Column(8)._WidthInPix=3598"
         Splits(1)._ColumnProps(67)=   "Column(8)._ColStyle=8708"
         Splits(1)._ColumnProps(68)=   "Column(8).Order=9"
         Splits(1)._ColumnProps(69)=   "Column(8)._MinWidth=79999936"
         Splits(1)._ColumnProps(70)=   "Column(9).Width=2170"
         Splits(1)._ColumnProps(71)=   "Column(9).DividerColor=0"
         Splits(1)._ColumnProps(72)=   "Column(9)._WidthInPix=2090"
         Splits(1)._ColumnProps(73)=   "Column(9)._ColStyle=8708"
         Splits(1)._ColumnProps(74)=   "Column(9).Order=10"
         Splits(1)._ColumnProps(75)=   "Column(9)._MinWidth=80007280"
         Splits(1)._ColumnProps(76)=   "Column(10).Width=1191"
         Splits(1)._ColumnProps(77)=   "Column(10).DividerColor=0"
         Splits(1)._ColumnProps(78)=   "Column(10)._WidthInPix=1111"
         Splits(1)._ColumnProps(79)=   "Column(10)._ColStyle=513"
         Splits(1)._ColumnProps(80)=   "Column(10).Order=11"
         Splits(1)._ColumnProps(81)=   "Column(10)._MinWidth=80007280"
         Splits(1)._ColumnProps(82)=   "Column(11).Width=4233"
         Splits(1)._ColumnProps(83)=   "Column(11).DividerColor=0"
         Splits(1)._ColumnProps(84)=   "Column(11)._WidthInPix=4154"
         Splits(1)._ColumnProps(85)=   "Column(11)._ColStyle=516"
         Splits(1)._ColumnProps(86)=   "Column(11).Order=12"
         Splits(1)._ColumnProps(87)=   "Column(11)._MinWidth=80007280"
         Splits(1)._ColumnProps(88)=   "Column(12).Width=2064"
         Splits(1)._ColumnProps(89)=   "Column(12).DividerColor=0"
         Splits(1)._ColumnProps(90)=   "Column(12)._WidthInPix=1984"
         Splits(1)._ColumnProps(91)=   "Column(12)._ColStyle=8705"
         Splits(1)._ColumnProps(92)=   "Column(12).Order=13"
         Splits(1)._ColumnProps(93)=   "Column(12)._MinWidth=80007280"
         Splits(1)._ColumnProps(94)=   "Column(13).Width=3016"
         Splits(1)._ColumnProps(95)=   "Column(13).DividerColor=0"
         Splits(1)._ColumnProps(96)=   "Column(13)._WidthInPix=2937"
         Splits(1)._ColumnProps(97)=   "Column(13)._ColStyle=8708"
         Splits(1)._ColumnProps(98)=   "Column(13).Order=14"
         Splits(1)._ColumnProps(99)=   "Column(13)._MinWidth=80010048"
         Splits(1)._ColumnProps(100)=   "Column(14).Width=2037"
         Splits(1)._ColumnProps(101)=   "Column(14).DividerColor=0"
         Splits(1)._ColumnProps(102)=   "Column(14)._WidthInPix=1958"
         Splits(1)._ColumnProps(103)=   "Column(14)._ColStyle=8705"
         Splits(1)._ColumnProps(104)=   "Column(14).Order=15"
         Splits(1)._ColumnProps(105)=   "Column(15).Width=2725"
         Splits(1)._ColumnProps(106)=   "Column(15).DividerColor=0"
         Splits(1)._ColumnProps(107)=   "Column(15)._WidthInPix=2646"
         Splits(1)._ColumnProps(108)=   "Column(15)._ColStyle=8705"
         Splits(1)._ColumnProps(109)=   "Column(15).Order=16"
         Splits(1)._ColumnProps(110)=   "Column(16).Width=2725"
         Splits(1)._ColumnProps(111)=   "Column(16).DividerColor=0"
         Splits(1)._ColumnProps(112)=   "Column(16)._WidthInPix=2646"
         Splits(1)._ColumnProps(113)=   "Column(16)._ColStyle=8708"
         Splits(1)._ColumnProps(114)=   "Column(16).Order=17"
         Splits(1)._ColumnProps(115)=   "Column(17).Width=2725"
         Splits(1)._ColumnProps(116)=   "Column(17).DividerColor=0"
         Splits(1)._ColumnProps(117)=   "Column(17)._WidthInPix=2646"
         Splits(1)._ColumnProps(118)=   "Column(17)._ColStyle=8708"
         Splits(1)._ColumnProps(119)=   "Column(17).Order=18"
         Splits(1)._ColumnProps(120)=   "Column(18).Width=2725"
         Splits(1)._ColumnProps(121)=   "Column(18).DividerColor=0"
         Splits(1)._ColumnProps(122)=   "Column(18)._WidthInPix=2646"
         Splits(1)._ColumnProps(123)=   "Column(18)._ColStyle=8708"
         Splits(1)._ColumnProps(124)=   "Column(18).Order=19"
         Splits(1)._ColumnProps(125)=   "Column(19).Width=2725"
         Splits(1)._ColumnProps(126)=   "Column(19).DividerColor=0"
         Splits(1)._ColumnProps(127)=   "Column(19)._WidthInPix=2646"
         Splits(1)._ColumnProps(128)=   "Column(19)._ColStyle=8705"
         Splits(1)._ColumnProps(129)=   "Column(19).Order=20"
         Splits(1)._ColumnProps(130)=   "Column(20).Width=2725"
         Splits(1)._ColumnProps(131)=   "Column(20).DividerColor=0"
         Splits(1)._ColumnProps(132)=   "Column(20)._WidthInPix=2646"
         Splits(1)._ColumnProps(133)=   "Column(20)._ColStyle=8708"
         Splits(1)._ColumnProps(134)=   "Column(20).Order=21"
         Splits(1)._ColumnProps(135)=   "Column(21).Width=2725"
         Splits(1)._ColumnProps(136)=   "Column(21).DividerColor=0"
         Splits(1)._ColumnProps(137)=   "Column(21)._WidthInPix=2646"
         Splits(1)._ColumnProps(138)=   "Column(21)._ColStyle=8708"
         Splits(1)._ColumnProps(139)=   "Column(21).Order=22"
         Splits(1)._ColumnProps(140)=   "Column(22).Width=2725"
         Splits(1)._ColumnProps(141)=   "Column(22).DividerColor=0"
         Splits(1)._ColumnProps(142)=   "Column(22)._WidthInPix=2646"
         Splits(1)._ColumnProps(143)=   "Column(22)._ColStyle=8705"
         Splits(1)._ColumnProps(144)=   "Column(22).Order=23"
         Splits(1)._ColumnProps(145)=   "Column(23).Width=2725"
         Splits(1)._ColumnProps(146)=   "Column(23).DividerColor=0"
         Splits(1)._ColumnProps(147)=   "Column(23)._WidthInPix=2646"
         Splits(1)._ColumnProps(148)=   "Column(23)._ColStyle=8708"
         Splits(1)._ColumnProps(149)=   "Column(23).Order=24"
         Splits(1)._ColumnProps(150)=   "Column(23)._MinWidth=80015760"
         Splits(1)._ColumnProps(151)=   "Column(24).Width=2725"
         Splits(1)._ColumnProps(152)=   "Column(24).DividerColor=0"
         Splits(1)._ColumnProps(153)=   "Column(24)._WidthInPix=2646"
         Splits(1)._ColumnProps(154)=   "Column(24)._ColStyle=516"
         Splits(1)._ColumnProps(155)=   "Column(24).Visible=0"
         Splits(1)._ColumnProps(156)=   "Column(24).Order=25"
         Splits.Count    =   2
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "DAFTAR KARYAWAN"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=90,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=86,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=82,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=78,.parent=13,.locked=-1"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=74,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.locked=-1"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=222,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=219,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=220,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=221,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=32,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=98,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=95,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=96,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=97,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=234,.parent=13,.alignment=2"
         _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=231,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=232,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=233,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=242,.parent=13"
         _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=239,.parent=14"
         _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=240,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=241,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=50,.parent=13"
         _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=47,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=48,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=49,.parent=17"
         _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=54,.parent=13"
         _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=51,.parent=14"
         _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=52,.parent=15"
         _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=53,.parent=17"
         _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=62,.parent=13"
         _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=59,.parent=14"
         _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=60,.parent=15"
         _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=61,.parent=17"
         _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=66,.parent=13"
         _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=63,.parent=14"
         _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=64,.parent=15"
         _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=65,.parent=17"
         _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=102,.parent=13"
         _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
         _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
         _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
         _StyleDefs(102) =   "Splits(0).Columns(17).Style:id=110,.parent=13"
         _StyleDefs(103) =   "Splits(0).Columns(17).HeadingStyle:id=107,.parent=14"
         _StyleDefs(104) =   "Splits(0).Columns(17).FooterStyle:id=108,.parent=15"
         _StyleDefs(105) =   "Splits(0).Columns(17).EditorStyle:id=109,.parent=17"
         _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=46,.parent=13"
         _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=43,.parent=14"
         _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=44,.parent=15"
         _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=45,.parent=17"
         _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=58,.parent=13"
         _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=55,.parent=14"
         _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=56,.parent=15"
         _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=57,.parent=17"
         _StyleDefs(114) =   "Splits(0).Columns(20).Style:id=94,.parent=13"
         _StyleDefs(115) =   "Splits(0).Columns(20).HeadingStyle:id=91,.parent=14"
         _StyleDefs(116) =   "Splits(0).Columns(20).FooterStyle:id=92,.parent=15"
         _StyleDefs(117) =   "Splits(0).Columns(20).EditorStyle:id=93,.parent=17"
         _StyleDefs(118) =   "Splits(0).Columns(21).Style:id=106,.parent=13"
         _StyleDefs(119) =   "Splits(0).Columns(21).HeadingStyle:id=103,.parent=14"
         _StyleDefs(120) =   "Splits(0).Columns(21).FooterStyle:id=104,.parent=15"
         _StyleDefs(121) =   "Splits(0).Columns(21).EditorStyle:id=105,.parent=17"
         _StyleDefs(122) =   "Splits(0).Columns(22).Style:id=118,.parent=13"
         _StyleDefs(123) =   "Splits(0).Columns(22).HeadingStyle:id=115,.parent=14"
         _StyleDefs(124) =   "Splits(0).Columns(22).FooterStyle:id=116,.parent=15"
         _StyleDefs(125) =   "Splits(0).Columns(22).EditorStyle:id=117,.parent=17"
         _StyleDefs(126) =   "Splits(0).Columns(23).Style:id=122,.parent=13"
         _StyleDefs(127) =   "Splits(0).Columns(23).HeadingStyle:id=119,.parent=14"
         _StyleDefs(128) =   "Splits(0).Columns(23).FooterStyle:id=120,.parent=15"
         _StyleDefs(129) =   "Splits(0).Columns(23).EditorStyle:id=121,.parent=17"
         _StyleDefs(130) =   "Splits(0).Columns(24).Style:id=114,.parent=13"
         _StyleDefs(131) =   "Splits(0).Columns(24).HeadingStyle:id=111,.parent=14"
         _StyleDefs(132) =   "Splits(0).Columns(24).FooterStyle:id=112,.parent=15"
         _StyleDefs(133) =   "Splits(0).Columns(24).EditorStyle:id=113,.parent=17"
         _StyleDefs(134) =   "Splits(1).Style:id=123,.parent=1"
         _StyleDefs(135) =   "Splits(1).CaptionStyle:id=132,.parent=4,.bgcolor=&H80000002&"
         _StyleDefs(136) =   ":id=132,.fgcolor=&H80000009&"
         _StyleDefs(137) =   "Splits(1).HeadingStyle:id=124,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
         _StyleDefs(138) =   ":id=124,.fgcolor=&H80000002&"
         _StyleDefs(139) =   "Splits(1).FooterStyle:id=125,.parent=3"
         _StyleDefs(140) =   "Splits(1).InactiveStyle:id=126,.parent=5"
         _StyleDefs(141) =   "Splits(1).SelectedStyle:id=128,.parent=6"
         _StyleDefs(142) =   "Splits(1).EditorStyle:id=127,.parent=7"
         _StyleDefs(143) =   "Splits(1).HighlightRowStyle:id=129,.parent=8"
         _StyleDefs(144) =   "Splits(1).EvenRowStyle:id=130,.parent=9"
         _StyleDefs(145) =   "Splits(1).OddRowStyle:id=131,.parent=10"
         _StyleDefs(146) =   "Splits(1).RecordSelectorStyle:id=133,.parent=11"
         _StyleDefs(147) =   "Splits(1).FilterBarStyle:id=134,.parent=12"
         _StyleDefs(148) =   "Splits(1).Columns(0).Style:id=138,.parent=123"
         _StyleDefs(149) =   "Splits(1).Columns(0).HeadingStyle:id=135,.parent=124"
         _StyleDefs(150) =   "Splits(1).Columns(0).FooterStyle:id=136,.parent=125"
         _StyleDefs(151) =   "Splits(1).Columns(0).EditorStyle:id=137,.parent=127"
         _StyleDefs(152) =   "Splits(1).Columns(1).Style:id=142,.parent=123"
         _StyleDefs(153) =   "Splits(1).Columns(1).HeadingStyle:id=139,.parent=124"
         _StyleDefs(154) =   "Splits(1).Columns(1).FooterStyle:id=140,.parent=125"
         _StyleDefs(155) =   "Splits(1).Columns(1).EditorStyle:id=141,.parent=127"
         _StyleDefs(156) =   "Splits(1).Columns(2).Style:id=146,.parent=123"
         _StyleDefs(157) =   "Splits(1).Columns(2).HeadingStyle:id=143,.parent=124"
         _StyleDefs(158) =   "Splits(1).Columns(2).FooterStyle:id=144,.parent=125"
         _StyleDefs(159) =   "Splits(1).Columns(2).EditorStyle:id=145,.parent=127"
         _StyleDefs(160) =   "Splits(1).Columns(3).Style:id=150,.parent=123"
         _StyleDefs(161) =   "Splits(1).Columns(3).HeadingStyle:id=147,.parent=124"
         _StyleDefs(162) =   "Splits(1).Columns(3).FooterStyle:id=148,.parent=125"
         _StyleDefs(163) =   "Splits(1).Columns(3).EditorStyle:id=149,.parent=127"
         _StyleDefs(164) =   "Splits(1).Columns(4).Style:id=154,.parent=123"
         _StyleDefs(165) =   "Splits(1).Columns(4).HeadingStyle:id=151,.parent=124"
         _StyleDefs(166) =   "Splits(1).Columns(4).FooterStyle:id=152,.parent=125"
         _StyleDefs(167) =   "Splits(1).Columns(4).EditorStyle:id=153,.parent=127"
         _StyleDefs(168) =   "Splits(1).Columns(5).Style:id=158,.parent=123"
         _StyleDefs(169) =   "Splits(1).Columns(5).HeadingStyle:id=155,.parent=124"
         _StyleDefs(170) =   "Splits(1).Columns(5).FooterStyle:id=156,.parent=125"
         _StyleDefs(171) =   "Splits(1).Columns(5).EditorStyle:id=157,.parent=127"
         _StyleDefs(172) =   "Splits(1).Columns(6).Style:id=226,.parent=123"
         _StyleDefs(173) =   "Splits(1).Columns(6).HeadingStyle:id=223,.parent=124"
         _StyleDefs(174) =   "Splits(1).Columns(6).FooterStyle:id=224,.parent=125"
         _StyleDefs(175) =   "Splits(1).Columns(6).EditorStyle:id=225,.parent=127"
         _StyleDefs(176) =   "Splits(1).Columns(7).Style:id=162,.parent=123,.locked=-1"
         _StyleDefs(177) =   "Splits(1).Columns(7).HeadingStyle:id=159,.parent=124"
         _StyleDefs(178) =   "Splits(1).Columns(7).FooterStyle:id=160,.parent=125"
         _StyleDefs(179) =   "Splits(1).Columns(7).EditorStyle:id=161,.parent=127"
         _StyleDefs(180) =   "Splits(1).Columns(8).Style:id=166,.parent=123,.locked=-1"
         _StyleDefs(181) =   "Splits(1).Columns(8).HeadingStyle:id=163,.parent=124"
         _StyleDefs(182) =   "Splits(1).Columns(8).FooterStyle:id=164,.parent=125"
         _StyleDefs(183) =   "Splits(1).Columns(8).EditorStyle:id=165,.parent=127"
         _StyleDefs(184) =   "Splits(1).Columns(9).Style:id=230,.parent=123,.locked=-1"
         _StyleDefs(185) =   "Splits(1).Columns(9).HeadingStyle:id=227,.parent=124"
         _StyleDefs(186) =   "Splits(1).Columns(9).FooterStyle:id=228,.parent=125"
         _StyleDefs(187) =   "Splits(1).Columns(9).EditorStyle:id=229,.parent=127"
         _StyleDefs(188) =   "Splits(1).Columns(10).Style:id=238,.parent=123,.alignment=2"
         _StyleDefs(189) =   "Splits(1).Columns(10).HeadingStyle:id=235,.parent=124"
         _StyleDefs(190) =   "Splits(1).Columns(10).FooterStyle:id=236,.parent=125"
         _StyleDefs(191) =   "Splits(1).Columns(10).EditorStyle:id=237,.parent=127"
         _StyleDefs(192) =   "Splits(1).Columns(11).Style:id=246,.parent=123"
         _StyleDefs(193) =   "Splits(1).Columns(11).HeadingStyle:id=243,.parent=124"
         _StyleDefs(194) =   "Splits(1).Columns(11).FooterStyle:id=244,.parent=125"
         _StyleDefs(195) =   "Splits(1).Columns(11).EditorStyle:id=245,.parent=127"
         _StyleDefs(196) =   "Splits(1).Columns(12).Style:id=170,.parent=123,.alignment=2,.locked=-1"
         _StyleDefs(197) =   "Splits(1).Columns(12).HeadingStyle:id=167,.parent=124"
         _StyleDefs(198) =   "Splits(1).Columns(12).FooterStyle:id=168,.parent=125"
         _StyleDefs(199) =   "Splits(1).Columns(12).EditorStyle:id=169,.parent=127"
         _StyleDefs(200) =   "Splits(1).Columns(13).Style:id=174,.parent=123,.locked=-1"
         _StyleDefs(201) =   "Splits(1).Columns(13).HeadingStyle:id=171,.parent=124"
         _StyleDefs(202) =   "Splits(1).Columns(13).FooterStyle:id=172,.parent=125"
         _StyleDefs(203) =   "Splits(1).Columns(13).EditorStyle:id=173,.parent=127"
         _StyleDefs(204) =   "Splits(1).Columns(14).Style:id=178,.parent=123,.alignment=2,.locked=-1"
         _StyleDefs(205) =   "Splits(1).Columns(14).HeadingStyle:id=175,.parent=124"
         _StyleDefs(206) =   "Splits(1).Columns(14).FooterStyle:id=176,.parent=125"
         _StyleDefs(207) =   "Splits(1).Columns(14).EditorStyle:id=177,.parent=127"
         _StyleDefs(208) =   "Splits(1).Columns(15).Style:id=182,.parent=123,.alignment=2,.locked=-1"
         _StyleDefs(209) =   "Splits(1).Columns(15).HeadingStyle:id=179,.parent=124"
         _StyleDefs(210) =   "Splits(1).Columns(15).FooterStyle:id=180,.parent=125"
         _StyleDefs(211) =   "Splits(1).Columns(15).EditorStyle:id=181,.parent=127"
         _StyleDefs(212) =   "Splits(1).Columns(16).Style:id=186,.parent=123,.locked=-1"
         _StyleDefs(213) =   "Splits(1).Columns(16).HeadingStyle:id=183,.parent=124"
         _StyleDefs(214) =   "Splits(1).Columns(16).FooterStyle:id=184,.parent=125"
         _StyleDefs(215) =   "Splits(1).Columns(16).EditorStyle:id=185,.parent=127"
         _StyleDefs(216) =   "Splits(1).Columns(17).Style:id=190,.parent=123,.locked=-1"
         _StyleDefs(217) =   "Splits(1).Columns(17).HeadingStyle:id=187,.parent=124"
         _StyleDefs(218) =   "Splits(1).Columns(17).FooterStyle:id=188,.parent=125"
         _StyleDefs(219) =   "Splits(1).Columns(17).EditorStyle:id=189,.parent=127"
         _StyleDefs(220) =   "Splits(1).Columns(18).Style:id=194,.parent=123,.locked=-1"
         _StyleDefs(221) =   "Splits(1).Columns(18).HeadingStyle:id=191,.parent=124"
         _StyleDefs(222) =   "Splits(1).Columns(18).FooterStyle:id=192,.parent=125"
         _StyleDefs(223) =   "Splits(1).Columns(18).EditorStyle:id=193,.parent=127"
         _StyleDefs(224) =   "Splits(1).Columns(19).Style:id=198,.parent=123,.alignment=2,.locked=-1"
         _StyleDefs(225) =   "Splits(1).Columns(19).HeadingStyle:id=195,.parent=124"
         _StyleDefs(226) =   "Splits(1).Columns(19).FooterStyle:id=196,.parent=125"
         _StyleDefs(227) =   "Splits(1).Columns(19).EditorStyle:id=197,.parent=127"
         _StyleDefs(228) =   "Splits(1).Columns(20).Style:id=202,.parent=123,.locked=-1"
         _StyleDefs(229) =   "Splits(1).Columns(20).HeadingStyle:id=199,.parent=124"
         _StyleDefs(230) =   "Splits(1).Columns(20).FooterStyle:id=200,.parent=125"
         _StyleDefs(231) =   "Splits(1).Columns(20).EditorStyle:id=201,.parent=127"
         _StyleDefs(232) =   "Splits(1).Columns(21).Style:id=206,.parent=123,.locked=-1"
         _StyleDefs(233) =   "Splits(1).Columns(21).HeadingStyle:id=203,.parent=124"
         _StyleDefs(234) =   "Splits(1).Columns(21).FooterStyle:id=204,.parent=125"
         _StyleDefs(235) =   "Splits(1).Columns(21).EditorStyle:id=205,.parent=127"
         _StyleDefs(236) =   "Splits(1).Columns(22).Style:id=214,.parent=123,.alignment=2,.locked=-1"
         _StyleDefs(237) =   "Splits(1).Columns(22).HeadingStyle:id=211,.parent=124"
         _StyleDefs(238) =   "Splits(1).Columns(22).FooterStyle:id=212,.parent=125"
         _StyleDefs(239) =   "Splits(1).Columns(22).EditorStyle:id=213,.parent=127"
         _StyleDefs(240) =   "Splits(1).Columns(23).Style:id=218,.parent=123,.locked=-1"
         _StyleDefs(241) =   "Splits(1).Columns(23).HeadingStyle:id=215,.parent=124"
         _StyleDefs(242) =   "Splits(1).Columns(23).FooterStyle:id=216,.parent=125"
         _StyleDefs(243) =   "Splits(1).Columns(23).EditorStyle:id=217,.parent=127"
         _StyleDefs(244) =   "Splits(1).Columns(24).Style:id=210,.parent=123"
         _StyleDefs(245) =   "Splits(1).Columns(24).HeadingStyle:id=207,.parent=124"
         _StyleDefs(246) =   "Splits(1).Columns(24).FooterStyle:id=208,.parent=125"
         _StyleDefs(247) =   "Splits(1).Columns(24).EditorStyle:id=209,.parent=127"
         _StyleDefs(248) =   "Named:id=33:Normal"
         _StyleDefs(249) =   ":id=33,.parent=0"
         _StyleDefs(250) =   "Named:id=34:Heading"
         _StyleDefs(251) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(252) =   ":id=34,.wraptext=-1"
         _StyleDefs(253) =   "Named:id=35:Footing"
         _StyleDefs(254) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(255) =   "Named:id=36:Selected"
         _StyleDefs(256) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(257) =   "Named:id=37:Caption"
         _StyleDefs(258) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(259) =   "Named:id=38:HighlightRow"
         _StyleDefs(260) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(261) =   "Named:id=39:EvenRow"
         _StyleDefs(262) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(263) =   "Named:id=40:OddRow"
         _StyleDefs(264) =   ":id=40,.parent=33"
         _StyleDefs(265) =   "Named:id=41:RecordSelector"
         _StyleDefs(266) =   ":id=41,.parent=34"
         _StyleDefs(267) =   "Named:id=42:FilterBar"
         _StyleDefs(268) =   ":id=42,.parent=33"
      End
      Begin VB.Label lbl_kontrak 
         Caption         =   "Nama Karyawan : ABDUL ROUF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74730
         TabIndex        =   268
         Top             =   390
         Width           =   5685
      End
      Begin VB.Label lbl_training 
         Caption         =   "Nama Karyawan : ABDUL ROUF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74730
         TabIndex        =   267
         Top             =   450
         Width           =   5685
      End
      Begin VB.Label lbl_grade 
         Caption         =   "Nama Karyawan : ABDUL ROUF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74700
         TabIndex        =   266
         Top             =   360
         Width           =   5685
      End
      Begin VB.Label lbl_jbtn 
         Caption         =   "Nama Karyawan : ABDUL ROUF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74760
         TabIndex        =   265
         Top             =   390
         Width           =   5685
      End
      Begin VB.Label lbl_exp 
         Caption         =   "Nama Karyawan : ABDUL ROUF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74700
         TabIndex        =   264
         Top             =   390
         Width           =   5685
      End
      Begin VB.Label lbl_skill 
         Caption         =   "Nama Karyawan : ABDUL ROUF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74730
         TabIndex        =   263
         Top             =   390
         Width           =   5685
      End
      Begin VB.Label lbl_edu 
         Caption         =   "Nama Karyawan : ABDUL ROUF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74730
         TabIndex        =   262
         Top             =   420
         Width           =   5685
      End
      Begin VB.Label lbl_fam 
         Caption         =   "Nama Karyawan : ABDUL ROUF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74730
         TabIndex        =   261
         Top             =   390
         Width           =   5685
      End
      Begin VB.Label lbl_employee 
         Caption         =   "Total Karyawan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   270
         TabIndex        =   77
         Top             =   750
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "PERUSAHAN"
         Height          =   195
         Left            =   270
         TabIndex        =   72
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER KARYAWAN"
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
      TabIndex        =   45
      Top             =   150
      Width           =   4365
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frmMstEmployee.frx":45455
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14850
   End
End
Attribute VB_Name = "frm_mst_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Const FO_DELETE = &H3
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim rsEmployee As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsDepartment As New ADODB.Recordset
Dim rsDivision As New ADODB.Recordset
Dim rsTitle As New ADODB.Recordset
Dim rsLevel As New ADODB.Recordset
Dim rsGrade As New ADODB.Recordset
Dim rsStatus As New ADODB.Recordset
Dim rsBank As New ADODB.Recordset

Dim rsFams As New ADODB.Recordset
Dim rsEdu As New ADODB.Recordset
Dim rsSkill As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsTraining As New ADODB.Recordset
Dim rsTDBTitle  As New ADODB.Recordset
Dim rsTitleEmp As New ADODB.Recordset
Dim rsTDBGrade As New ADODB.Recordset
Dim rsGrade1 As New ADODB.Recordset
Dim rsContract As New ADODB.Recordset

Dim clsReg As New clsCheckRegister

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim i_lang As Integer
Dim v_company As String
Dim dst As String
Dim v_empCode As String
Dim v_seq As Integer
Dim v_flag_periode As Integer
Dim src As String

Private Function check_validate_exist_new() As Boolean
Dim tanya As Integer
    check_validate_exist_new = False
    
    If rs.State Then rs.Close
        If SSTab1.Tab = 0 Then
            SQL = "select count(employee_code) as rec_count from m_employee where employee_code = '" _
                        & Replace$(Trim$(txt_employee_code), Chr$(39), Chr$(96)) & "' or nik = '" & Trim(txt_nik.Text) & "'"
        rs.Open SQL, CnG, adOpenStatic, adLockReadOnly
        
            If rs.Fields("rec_count").Value > 0 Then
                If SSTab1.Tab = 0 Then
                    tanya = MsgBox("Data Sudah Ada.." & Chr(13) & "Apakah Ingin Memasukkan Data Karyawan di Perusahaan Ini?", vbExclamation + vbYesNo, "Informasi!")
                    If tanya = vbYes Then
                        check_validate_exist_new = False
                    Else
                        check_validate_exist_new = True
                    End If
                Else
                    check_validate_exist_new = True
                End If
                rs.Close
                Exit Function
            End If
        rs.Close
    End If
End Function

Private Sub check_invalid()
    If SSTab1.Tab = 0 Then
        MsgBox "Data Sudah Ada...", vbCritical, headerMSG
        txt_employee_code = ""
        If txt_nik.Enabled = True Then txt_nik.SetFocus
    End If
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If check_validate_exist_new Then
        check_validate_exist_edit = True
        Exit Function
    End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True
    
    If SSTab1.Tab = 0 Then
        'validasi nik
        If Trim(txt_nik.Text) = "" Then
            MsgBox "Kode Karyawan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_nik.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi employee name
        If Trim(txt_employee_name.Text) = "" Then
            MsgBox "Nama Karyawan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_employee_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi employee nick name
        If Trim(txt_employee_nick_name.Text) = "" Then
            MsgBox "Nama Pendek Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_employee_nick_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi division tdbcombo
        If check_validate_tdbcombo(TDBCombo_division) = False Then
            MsgBox "Divisi Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_division.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
'        'validasi title
'        If check_validate_tdbcombo(TDBCombo_title) = False Then
'            MsgBox "Title Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
'            TDBCombo_title.SetFocus
'            check_validate_new = False
'            Exit Function
'        End If
        
        'validasi level
        If check_validate_tdbcombo(TDBCombo_level) = False Then
            MsgBox "Level Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_level.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi grade
        If check_validate_tdbcombo(TDBCombo_grade) = False Then
            MsgBox "Grade Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_grade.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi status
        If check_validate_tdbcombo(TDBCombo_status) = False Then
            MsgBox "Status Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_status.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi place of birth
        If txt_place_of_birth.Text = "" Then
            MsgBox "Tempat Lahir Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_place_of_birth.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi sex
        If cbo_sex.Text = "" Then
            MsgBox "Jenis Kelamin Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            cbo_sex.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi religion
        If cbo_religion.Text = "" Then
            MsgBox "Agama Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            cbo_religion.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi place of birth
        If txt_number_of_children.Text = "" Then
            MsgBox "Jumlah Anak Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_number_of_children.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        'validasi name
        If txt_family_name.Text = "" Then
            MsgBox "Nama Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_family_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi relationship
        If cbo_fams_rel.Text = "" Then
            MsgBox "Hubungan Keluarga Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            cbo_fams_rel.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi sex
        If cbo_fams_sex.Text = "" Then
            MsgBox "Jenis Kelamin Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            cbo_fams_sex.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        'validasi sex
        If cbo_edu_level.Text = "" Then
            MsgBox "Level Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            cbo_edu_level.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi name
        If txt_edu_school.Text = "" Then
            MsgBox "Sekolah/Universitas Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_edu_school.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 3 Then
        'validasi skill level
        If cbo_skill_level.Text = "" Then
            MsgBox "Level Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            cbo_skill_level.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi name
        If txt_skill_name.Text = "" Then
            MsgBox "Nama KEahlian Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_skill_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 4 Then
        'validasi company name
        If txt_job_company.Text = "" Then
            MsgBox "Nama Perusahaan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_job_company.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi last salary
        If Not IsNumeric(txt_job_salary.Text) Or Val(DropAllComma(txt_job_salary.Text)) < 0 Then
            MsgBox "Nilai Upah Terakhir Tidak Sesuai", vbOKOnly + vbInformation, headerMSG
            txt_job_salary.Text = 0
            txt_job_salary.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 5 Then
        'validasi title
        If check_validate_tdbcombo(TDBCombo_title_emp) = False Then
            MsgBox "Jabatan Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_title_emp.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 6 Then
        'validasi grade
        If check_validate_tdbcombo(TDBCombo_grade1) = False Then
            MsgBox "Grade Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_grade1.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 7 Then
        'validasi training subject
        If txt_training_subject.Text = "" Then
            MsgBox "Nama Training Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_training_subject.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi organizer
        If txt_training_organize.Text = "" Then
            MsgBox "Penyelenggara Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_training_organize.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi training place
        If txt_training_place.Text = "" Then
            MsgBox "Tempat Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_training_place.SetFocus
            check_validate_new = False
            Exit Function
        End If
    
    ElseIf SSTab1.Tab = 8 Then
        'validasi no contract
        If txt_contract_no.Text = "" Then
            MsgBox "No. Kontrak Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_contract_no.SetFocus
            check_validate_new = False
            Exit Function
        End If
    End If

End Function

'Private Sub cmd_import_Click()
'    'If check_validate_tdbcombo(TDBCombo_company) = False Then
'    '    MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
'    '    Exit Sub
'    'End If
'
'    frm_trans_import_employee.Show 1
'End Sub

Private Sub cancel_data()
    If SSTab1.Tab = 0 Then
        fra_status_emp.Visible = True
    End If
    
    int_mode = 0
    Call load_mode
End Sub

Private Sub delete_data()
Dim i As Integer
Dim pict As String

On Error GoTo Err
    If SSTab1.Tab = 0 Then
        If Not (TDBGrid_Employee.ApproxCount > 0 And TDBGrid_Employee.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Semua Informasi Mengenai Karyawan Ini Akan Dihapus..." & Chr(13) & _
                    "Apakah Yakin Akan Menghapus Data '" _
                    & TDBGrid_Employee.Columns("employee_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        
        CnG.Execute "DELETE FROM h_attendance where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM h_log_attendance where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM h_log_attendance_copy where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM h_log_attendance_reproccess where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM h_salary where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM t_check where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM t_employee_perf where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM t_employee_perf_result where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM t_income_expense where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM t_leave where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM t_leave_periode where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM t_pensiun where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM t_spl where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM t_spl_auto where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM t_thr where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM td_loan where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM td_shift where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM td_shift2 where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "DELETE FROM tm_loan where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        
        CnG.Execute "delete from m_employee where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "delete from m_employee_fams where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "delete from m_employee_edu where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "delete from m_employee_skill where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "delete from m_employee_exp where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "delete from m_employee_title where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "delete from m_employee_grade where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "delete from m_employee_training where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        CnG.Execute "delete from m_employee_contract where employee_code = '" _
            & TDBGrid_Employee.Columns("employee_code").Value & "'"
        
        CnG.CommitTrans
    
        Call load_data_employee
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_Family.ApproxCount > 0 And TDBGrid_Family.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Family.Columns("name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_employee_fams where employee_code = '" _
            & TDBGrid_Family.Columns("employee_code").Value & "' " & _
            "AND seq_no = '" & TDBGrid_Family.Columns("seq_no").Value & "'"
        CnG.CommitTrans
    
        Call load_data_employee_family
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_Education.ApproxCount > 0 And TDBGrid_Education.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Education.Columns("school").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_employee_edu where employee_code = '" _
            & TDBGrid_Education.Columns("employee_code").Value & "' " & _
            "AND seq_no = '" & TDBGrid_Education.Columns("seq_no").Value & "'"
        CnG.CommitTrans
    
        Call load_data_employee_education
    ElseIf SSTab1.Tab = 3 Then
        If Not (TDBGrid_Skill.ApproxCount > 0 And TDBGrid_Skill.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Skill.Columns("skill").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_employee_skill where employee_code = '" _
            & TDBGrid_Skill.Columns("employee_code").Value & "' " & _
            "AND seq_no = '" & TDBGrid_Skill.Columns("seq_no").Value & "'"
        CnG.CommitTrans
    
        Call load_data_employee_skill
    ElseIf SSTab1.Tab = 4 Then
        If Not (TDBGrid_Job.ApproxCount > 0 And TDBGrid_Job.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Job.Columns("company_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_employee_exp where employee_code = '" _
            & TDBGrid_Job.Columns("employee_code").Value & "' " & _
            "AND seq_no = '" & TDBGrid_Job.Columns("seq_no").Value & "'"
        CnG.CommitTrans
    
        Call load_data_employee_job
    ElseIf SSTab1.Tab = 5 Then
        If Not (TDBGrid_Title.ApproxCount > 0 And TDBGrid_Title.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Title.Columns("title_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_employee_title where employee_code = '" _
            & TDBGrid_Title.Columns("employee_code").Value & "' " & _
            "AND seq_no = '" & TDBGrid_Title.Columns("seq_no").Value & "'"
        CnG.CommitTrans
    
        Call load_data_employee_title
    ElseIf SSTab1.Tab = 6 Then
        If Not (TDBGrid_Grade.ApproxCount > 0 And TDBGrid_Grade.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Grade.Columns("grade_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_employee_grade where employee_code = '" _
            & TDBGrid_Grade.Columns("employee_code").Value & "' " & _
            "AND seq_no = '" & TDBGrid_Grade.Columns("seq_no").Value & "'"
        CnG.CommitTrans
    
        Call load_data_employee_grade
    ElseIf SSTab1.Tab = 7 Then
        If Not (TDBGrid_Training.ApproxCount > 0 And TDBGrid_Training.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Training.Columns("material").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_employee_training where employee_code = '" _
            & TDBGrid_Training.Columns("employee_code").Value & "' " & _
            "AND seq_no = '" & TDBGrid_Training.Columns("seq_no").Value & "'"
        CnG.CommitTrans
    
        Call load_data_employee_training
    ElseIf SSTab1.Tab = 8 Then
        If Not (TDBGrid_Contract.ApproxCount > 0 And TDBGrid_Contract.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Contract.Columns("no_contract").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_employee_contract where employee_code = '" _
            & TDBGrid_Contract.Columns("employee_code").Value & "' " & _
            "AND seq_no = '" & TDBGrid_Contract.Columns("seq_no").Value & "'"
        CnG.CommitTrans
    
        Call load_data_employee_contract
    End If
    
    int_mode = 0
    Call load_mode
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

'Private Sub set_data_department(ByVal str_code As String)
'On Error GoTo Err
'
'    rsDepartment.MoveFirst
'    rsDepartment.Find ("department_code='" & str_code & "'")   ', 0, adSearchForward, 1)
'    If Not (rsDepartment.EOF = True Or rsDepartment.BOF = True) Then
'        TDBCombo_department.Bookmark = rsDepartment.AbsolutePosition
'        Call TDBCombo_department_itemChange
'    Else
'        TDBCombo_department.Text = ""
'    End If
'    Exit Sub
'
'Err:
'MsgBox Err.Description, vbExclamation, headerMSG
'End Sub

Private Sub set_data_division(ByRef kunci As String)
On Error GoTo Err
    
    rsDivision.MoveFirst
    rsDivision.Find ("division_code='" & kunci & "'")   ', 0, adSearchForward, 1)
    If Not (rsDivision.EOF = True Or rsDivision.BOF = True) Then
        TDBCombo_division.Bookmark = rsDivision.AbsolutePosition
        Call tdbcombo_division_itemChange
    Else
        TDBCombo_division.Text = ""
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub set_data_Title(ByVal str_code As String)
On Error GoTo Err

    rsTitle.MoveFirst
    rsTitle.Find ("title_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (rsTitle.EOF = True Or rsTitle.BOF = True) Then
        TDBCombo_title.Bookmark = rsTitle.AbsolutePosition
        Call TDBCombo_title_ItemChange
    Else
        TDBCombo_title.Text = ""
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub set_data_title_emp(ByVal str_code As String)
On Error GoTo Err
    
    rsTitleEmp.MoveFirst
    rsTitleEmp.Find ("title_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (rsTitleEmp.EOF = True Or rsTitleEmp.BOF = True) Then
        TDBCombo_title_emp.Bookmark = rsTitleEmp.AbsolutePosition
        Call TDBCombo_title_emp_ItemChange
    Else
        TDBCombo_title_emp.Text = ""
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub set_data_level(ByVal str_code As String)
On Error GoTo Err
    
    rsLevel.MoveFirst
    rsLevel.Find ("level_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (rsLevel.EOF = True Or rsLevel.BOF = True) Then
        TDBCombo_level.Bookmark = rsLevel.AbsolutePosition
        Call TDBCombo_level_ItemChange
    Else
        TDBCombo_level.Text = ""
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub set_data_grade(ByVal str_code As String)
On Error GoTo Err
    
    rsGrade.MoveFirst
    rsGrade.Find ("grade_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (rsGrade.EOF = True Or rsGrade.BOF = True) Then
        TDBCombo_grade.Bookmark = rsGrade.AbsolutePosition
        Call TDBCombo_grade_ItemChange
    Else
        TDBCombo_grade.Text = ""
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub set_data_grade1(ByVal str_code As String)
On Error GoTo Err
    
    rsTDBGrade.MoveFirst
    rsTDBGrade.Find ("grade_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (rsTDBGrade.EOF = True Or rsTDBGrade.BOF = True) Then
        TDBCombo_grade1.Bookmark = rsTDBGrade.AbsolutePosition
        Call TDBCombo_grade1_ItemChange
    Else
        TDBCombo_grade1.Text = ""
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
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

Private Sub set_data_bank(ByVal str_code As String)
On Error Resume Next
    
    rsBank.MoveFirst
    rsBank.Find ("bank_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (rsBank.EOF = True Or rsBank.BOF = True) Then
        TDBCombo_bank.Bookmark = rsBank.AbsolutePosition
        Call TDBCombo_bank_ItemChange
    Else
        TDBCombo_bank.Text = ""
    End If
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
Dim v_flag_active As Integer
    vSetData = 1
    If SSTab1.Tab = 0 Then
        TDBCombo_title.Enabled = False
        TDBCombo_grade.Enabled = False
        
        If Not (TDBGrid_Employee.ApproxCount > 0 And TDBGrid_Employee.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsEmployee
            v_flag_active = .Fields("flag_active").Value
            Call load_data_division
            
            txt_nik.Text = .Fields("nik").Value
            txt_employee_code.Text = .Fields("employee_code").Value
            txt_employee_name.Text = .Fields("employee_name").Value
            txt_employee_nick_name.Text = "" & .Fields("employee_nick_name").Value
            
            SQL = "SELECT * FROM m_employee_title " & _
                  "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
                  "ORDER BY seq_no DESC LIMIT 1"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                Call set_data_Title(rs!title_code)
            End If
            rs.Close
            
'            Call set_data_department(.Fields("department_code").Value)
            Call set_data_division(.Fields("division_code").Value)
            Call set_data_level(.Fields("level_code").Value)
            
            SQL = "SELECT * FROM m_employee_grade " & _
                  "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
                  "ORDER BY seq_no DESC LIMIT 1"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                Call set_data_grade(rs!grade_code)
            End If
            rs.Close
            
            Call set_data_status(.Fields("status_code").Value)
            '----------------------------------------------------------------------
                        
            txt_place_of_birth.Text = get_data_obj(txt_place_of_birth, .Fields("status_code").Value)
            DTPicker_birth.Value = get_data_obj(DTPicker_birth, .Fields("date_birth").Value)
            cbo_sex.ListIndex = get_data_obj(cbo_sex, .Fields("sex").Value)
            cbo_religion.ListIndex = get_data_obj(cbo_religion, .Fields("religion").Value)
            cbo_marital_status.ListIndex = get_data_obj(cbo_marital_status, .Fields("marital_status").Value)
            
            DTPicker_birth.Value = get_data_obj(DTPicker_birth, .Fields("date_birth").Value)
            txt_place_of_birth = get_data_obj(txt_place_of_birth, .Fields("place_birth").Value)
            cbo_sex.ListIndex = get_data_obj(cbo_sex, .Fields("sex").Value)
            cbo_religion.ListIndex = get_data_obj(cbo_religion, .Fields("religion").Value)
            cbo_marital_status.ListIndex = get_data_obj(cbo_marital_status, .Fields("marital_status").Value)
            txt_number_of_children.Text = get_data_obj(txt_number_of_children, .Fields("no_of_children").Value)
            txt_npwp.Text = get_data_obj(txt_npwp, .Fields("npwp").Value)
            cbo_tax_method.ListIndex = get_data_obj(cbo_tax_method, .Fields("npwp_method").Value)
            DTPicker_tglnpwp.Value = get_data_obj(DTPicker_tglnpwp, .Fields("npwp_registered_date").Value)
            txtAlamat.Text = get_data_obj(txt_address, .Fields("emp_address").Value)
            txt_address.Text = get_data_obj(txt_address, .Fields("npwp_address").Value)
            txt_jamsostek.Text = get_data_obj(txt_jamsostek, .Fields("no_jamsostek").Value)
            DTPicker_jstk.Value = get_data_obj(DTPicker_jstk, .Fields("jstk_registered_date").Value)
            Call set_data_bank(IIf(IsNull(.Fields("bank_code").Value), "", .Fields("bank_code").Value))
            txt_bank_account.Text = get_data_obj(txt_bank_account, .Fields("bank_account").Value)
            txt_account_name.Text = get_data_obj(txt_account_name, .Fields("bank_acc_name").Value)
            '---------------------------------------------------------------------
            
            DTPicker_start_working.Value = get_data_obj(DTPicker_start_working, .Fields("start_working").Value)
            DTPicker_appointment.Value = get_data_obj(DTPicker_appointment, .Fields("appointment_date").Value)
            txt_phone_number.Text = get_data_obj(txt_phone_number, .Fields("phone_number").Value)
            txt_email.Text = get_data_obj(txt_email, .Fields("email").Value)
            cbo_nationality.ListIndex = get_data_obj(cbo_nationality, .Fields("nationality").Value)
            txtNoKTP.Text = get_data_obj(txtNoKTP, .Fields("identity_number").Value)
            txt_description.Text = get_data_obj(txt_description, .Fields("description").Value)
            opt_active.Value = IIf(v_flag_active = 1, 1, 0)
            opt_not_active.Value = IIf(v_flag_active = 0, 1, 0)
            If opt_not_active Then
                DTPicker_end_working.Value = get_data_obj(DTPicker_end_working, .Fields("end_working").Value)
                txt_end_working_reason = get_data_obj(txt_end_working_reason, .Fields("reason").Value)
            End If
            '---------------------------------------------------------------------
            
'            '------------------------ Show Picture -------------------------------
'            If .Fields("picture").Value <> Null Or .Fields("picture").Value <> "" Then
'                img.Picture = LoadPicture(get_data_obj(img, .Fields("picture").Value))
'            End If
'
'            img.Width = img.Picture.Width
'            img.Height = img.Picture.Height
'            If pic.Width < img.Width Then
'                img.Width = pic.Width
''                img.Height = img.Height / (img.Picture.Width / img.Width)
'            End If
'
'            If pic.Height < img.Height Then
'                img.Height = pic.Height
''                img.Width = img.Width / (img.Picture.Height / img.Height)
'            End If
'
'            img.Left = 0
'            img.Top = 0
'
'            txt_pict_location.Text = .Fields("picture").Value
'            '---------------------------------------------------------------------
            
            '------------------------------- Show Image ---------------------------
            SQL = "SELECT picture FROM m_employee WHERE employee_code = '" & .Fields("employee_code").Value & "'"
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
            
            SQL = "SELECT * FROM m_employee_edu " & _
                  "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
                  "ORDER BY seq_no DESC LIMIT 1"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                txt_last_edu.Text = rs!nm_jenjang & " - " & rs!jurusan & " - " & rs!school
            End If
            rs.Close
            
            SQL = "SELECT * FROM m_employee_exp " & _
                  "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
                  "ORDER BY seq_no DESC LIMIT 1"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                txt_last_emp.Text = rs!company_name & " - " & rs!last_title
            End If
            rs.Close
            
            Call set_age_data
            Call set_working_age_data
        End With
        
        v_company = TDBCombo_company.Text
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_Family.ApproxCount > 0 And TDBGrid_Family.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsFams
            txt_family_name.Text = .Fields("name").Value
            cbo_fams_rel.ListIndex = .Fields("relationship").Value
            DTPicker_fams_birth.Value = .Fields("date_birth").Value
            cbo_fams_sex.ListIndex = .Fields("sex").Value
            txt_fams_edu.Text = .Fields("education").Value
            txt_fams_employment.Text = .Fields("employment").Value
            chk_fams_address.Value = .Fields("chk_address").Value
            txt_fams_address.Text = .Fields("address").Value
        End With
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_Education.ApproxCount > 0 And TDBGrid_Education.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsEdu
            DTPicker_edu_start.Value = .Fields("start_year").Value
            DTPicker_edu_end.Value = .Fields("end_year").Value
            cbo_edu_level.ListIndex = .Fields("jenjang").Value
            txt_edu_majors.Text = .Fields("jurusan").Value
            txt_edu_school.Text = .Fields("school").Value
            txt_edu_city.Text = .Fields("city").Value
            txt_edu_country.Text = .Fields("country").Value
        End With
    ElseIf SSTab1.Tab = 3 Then
        If Not (TDBGrid_Skill.ApproxCount > 0 And TDBGrid_Skill.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsSkill
            txt_skill_name.Text = .Fields("skill").Value
            cbo_skill_level.ListIndex = .Fields("level").Value
            txt_skill_description.Text = .Fields("description").Value
        End With
    ElseIf SSTab1.Tab = 4 Then
        If Not (TDBGrid_Job.ApproxCount > 0 And TDBGrid_Job.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsJob
            DTPicker_job_start.Value = .Fields("start_working").Value
            DTPicker_job_end.Value = .Fields("end_working").Value
            txt_job_company.Text = .Fields("company_name").Value
            txt_job_line.Text = .Fields("usaha").Value
            txt_job_dept.Text = .Fields("department_name").Value
            txt_job_title.Text = .Fields("last_title").Value
            txt_job_salary.Text = FormatNumber(.Fields("last_salary").Value)
            txt_job_reason.Text = .Fields("reason_stop_working").Value
            txt_job_description.Text = .Fields("description").Value
        End With
    ElseIf SSTab1.Tab = 5 Then
        If Not (TDBGrid_Title.ApproxCount > 0 And TDBGrid_Title.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsTitleEmp
            DTPicker_title.Value = .Fields("date_title").Value
            Call set_data_title_emp(.Fields("title_code").Value)
            txt_description_title.Text = IIf(IsNull(.Fields("description").Value), "", .Fields("description").Value)
        End With
    ElseIf SSTab1.Tab = 6 Then
        If Not (TDBGrid_Grade.ApproxCount > 0 And TDBGrid_Grade.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsGrade1
            DTPicker_grade.Value = .Fields("date_grade").Value
            Call set_data_grade1(.Fields("grade_code").Value)
            txt_grade_description.Text = IIf(IsNull(.Fields("description").Value), "", .Fields("description").Value)
        End With
    ElseIf SSTab1.Tab = 7 Then
        If Not (TDBGrid_Training.ApproxCount > 0 And TDBGrid_Training.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsTraining
            DTPicker_training_start.Value = .Fields("start_date").Value
            DTPicker_training_end.Value = .Fields("end_date").Value
            txt_training_subject.Text = .Fields("material").Value
            txt_training_organize.Text = .Fields("organizer").Value
            txt_training_place.Text = .Fields("place").Value
            txt_training_value.Text = .Fields("value").Value
        End With
    ElseIf SSTab1.Tab = 8 Then
        If Not (TDBGrid_Contract.ApproxCount > 0 And TDBGrid_Contract.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsContract
            DTPicker_contract_start.Value = .Fields("start_date").Value
            DTPicker_contract_end.Value = .Fields("end_date").Value
            txt_contract_no.Text = .Fields("no_contract").Value
            txt_contract_description.Text = .Fields("description").Value
        End With
    End If
End Sub

Private Sub edit_data()
    If SSTab1.Tab = 0 Then
        fra_status_emp.Visible = False
    End If
        
    int_mode = 2
    Call load_mode
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub new_data()
    If SSTab1.Tab = 0 Then
        fra_status_emp.Visible = False
    End If
    
    int_mode = 1
    Call load_mode
End Sub

Private Sub insert_new_data()

'On Error GoTo Err
    If SSTab1.Tab = 0 Then
        
        If clsReg.masihTrial = True Then
            If clsReg.batasKaryawan() = True Then
                MsgBox "Aplikasi Berjalan Dalam Masa Trial!" & Chr(13) & _
                        "Data Karyawan Dibatasi...", vbCritical, headerMSG
                Exit Sub
            End If
        End If
        
        SQL = "SELECT MAX(seq_no) jmlSeq FROM m_employee WHERE nik = '" & txt_nik.Text & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            v_seq = IIf(IsNull(rs!jmlSeq), 0, rs!jmlSeq) + 1
        Else
            v_seq = 1
        End If
        rs.Close
        
        v_empCode = Replace(txt_nik.Text, " ", "") & v_seq
        
        CnG.BeginTrans
        
        SQL = "INSERT INTO m_employee_title (employee_code,seq_no,date_title,title_code,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & v_empCode & "',1," & _
                "'" & Format(DTPicker_start_working.Value, "yyyy-MM-dd") & "','" & TDBCombo_title.Text & "'," & _
                "now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
        
        SQL = "INSERT INTO m_employee_grade (employee_code,seq_no,date_grade,grade_code,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & v_empCode & "',1," & _
                "'" & Format(DTPicker_start_working.Value, "yyyy-MM-dd") & "','" & TDBCombo_grade.Text & "'," & _
                "now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
        
        SQL = "INSERT INTO m_employee(employee_code,seq_no,nik,employee_name,employee_nick_name," & _
                "company_code,division_code,title_code,level_code,grade_code,status_code," & _
                "place_birth,date_birth,sex,religion,marital_status,no_of_children,emp_address,npwp," & _
                "npwp_method,npwp_registered_date,npwp_address,no_jamsostek,jstk_registered_date,bank_code," & _
                "bank_account,bank_acc_name,start_working,appointment_date,phone_number,email,nationality,identity_number," & _
                "description,flag_active,end_working,reason,picture,entry_date,entry_user) " & _
               "VALUES( " & _
                "'" & v_empCode & "','" & v_seq & "','" & txt_nik.Text & "','" & Replace(txt_employee_name.Text, "'", "''") & "','" & Replace(txt_employee_nick_name.Text, "'", "''") & "'," & _
                "'" & TDBCombo_company.Text & "','" & TDBCombo_division.Text & "','" & TDBCombo_title.Text & "'," & _
                "'" & TDBCombo_level.Text & "','" & TDBCombo_grade.Text & "','" & TDBCombo_status.Text & "','" & txt_place_of_birth.Text & "'," & _
                "'" & Format(DTPicker_birth.Value, "yyyy-MM-dd") & "','" & cbo_sex.ListIndex & "','" & cbo_religion.ListIndex & "','" & cbo_marital_status.ListIndex & "'," & _
                "'" & txt_number_of_children.Text & "','" & Replace(txtAlamat.Text, "'", "''") & "','" & txt_npwp.Text & "','" & cbo_tax_method.ListIndex & "','" & Format(DTPicker_tglnpwp.Value, "yyyy-MM-dd") & "'," & _
                "'" & Replace(txt_address.Text, "'", "''") & "','" & txt_jamsostek.Text & "','" & Format(DTPicker_jstk.Value, "yyyy-MM-dd") & "','" & TDBCombo_bank.Text & "'," & _
                "'" & txt_bank_account.Text & "','" & Replace(txt_account_name.Text, "'", "''") & "','" & Format(DTPicker_start_working.Value, "yyyy-MM-dd") & "','" & Format(DTPicker_appointment.Value, "yyyy-MM-dd") & "'," & _
                "'" & txt_phone_number.Text & "','" & txt_email.Text & "','" & cbo_nationality.ListIndex & "','" & txtNoKTP.Text & "'," & _
                "'" & Replace(txt_description.Text, "'", "''") & "','" & IIf(opt_active, 1, 0) & "','" & IIf(opt_active, "0000-00-00", Format(DTPicker_end_working.Value, "yyyy-MM-dd")) & "'," & _
                "'" & IIf(opt_active, "", txt_end_working_reason.Text) & "','" & img & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
        
        If fileExists(src) Then
            SQL = "SELECT employee_code, picture FROM m_employee WHERE employee_code = '" & txt_nik.Text & v_seq & "'"
            If Not addImageToDB(SQL, src, "picture") Then MsgBox "Gambar Gagal Disimpan...", vbExclamation, headerMSG
        End If
        
    ElseIf SSTab1.Tab = 1 Then
        SQL = "SELECT MAX(seq_no) jmlSeq FROM m_employee_fams " & _
                "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            v_seq = IIf(IsNull(rs!jmlSeq), 0, rs!jmlSeq) + 1
        Else
            v_seq = 1
        End If
        rs.Close
        
        CnG.BeginTrans
        
        SQL = "INSERT INTO m_employee_fams (employee_code,seq_no,name,relationship,nm_rel," & _
                "date_birth,sex,education,employment,chk_address,address,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & TDBGrid_Employee.Columns("employee_code").Value & "','" & v_seq & "'," & _
                "'" & Replace(txt_family_name.Text, "'", "''") & "','" & cbo_fams_rel.ListIndex & "'," & _
                "'" & cbo_fams_rel.Text & "','" & Format(DTPicker_fams_birth.Value, "yyyy-MM-dd") & "'," & _
                "'" & cbo_fams_sex.ListIndex & "','" & txt_fams_edu.Text & "','" & txt_fams_employment.Text & "'," & _
                "'" & chk_fams_address.Value & "','" & txt_fams_address.Text & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "SELECT MAX(seq_no) jmlSeq FROM m_employee_edu " & _
                "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            v_seq = IIf(IsNull(rs!jmlSeq), 0, rs!jmlSeq) + 1
        Else
            v_seq = 1
        End If
        rs.Close
        
        CnG.BeginTrans
        
        SQL = "INSERT INTO m_employee_edu (employee_code,seq_no,start_year,end_year,jenjang," & _
                "nm_jenjang,jurusan,school,city,country,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & TDBGrid_Employee.Columns("employee_code").Value & "','" & v_seq & "'," & _
                "'" & Format(DTPicker_edu_start.Value, "yyyy-MM-dd") & "','" & Format(DTPicker_edu_end.Value, "yyyy-MM-dd") & "'," & _
                "'" & cbo_edu_level.ListIndex & "','" & cbo_edu_level.Text & "','" & txt_edu_majors.Text & "'," & _
                "'" & txt_edu_school.Text & "','" & txt_edu_city.Text & "','" & txt_edu_country.Text & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 3 Then
        SQL = "SELECT MAX(seq_no) jmlSeq FROM m_employee_skill " & _
                "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            v_seq = IIf(IsNull(rs!jmlSeq), 0, rs!jmlSeq) + 1
        Else
            v_seq = 1
        End If
        rs.Close
        
        CnG.BeginTrans
        
        SQL = "INSERT INTO m_employee_skill (employee_code,seq_no,skill,level,nm_level," & _
                "description,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & TDBGrid_Employee.Columns("employee_code").Value & "','" & v_seq & "'," & _
                "'" & Replace(txt_skill_name.Text, "'", "''") & "','" & cbo_skill_level.ListIndex & "'," & _
                "'" & cbo_skill_level.Text & "','" & txt_skill_description.Text & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 4 Then
        SQL = "SELECT MAX(seq_no) jmlSeq FROM m_employee_exp " & _
                "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            v_seq = IIf(IsNull(rs!jmlSeq), 0, rs!jmlSeq) + 1
        Else
            v_seq = 1
        End If
        rs.Close
        
        CnG.BeginTrans
        
        SQL = "INSERT INTO m_employee_exp (employee_code,seq_no,start_working,end_working,company_name," & _
                "usaha,department_name,last_title,last_salary,reason_stop_working,description,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & TDBGrid_Employee.Columns("employee_code").Value & "','" & v_seq & "'," & _
                "'" & Format(DTPicker_job_start.Value, "yyyy-MM-dd") & "','" & Format(DTPicker_job_end.Value, "yyyy-MM-dd") & "'," & _
                "'" & txt_job_company.Text & "','" & txt_job_line.Text & "','" & txt_job_dept.Text & "'," & _
                "'" & txt_job_title.Text & "','" & Val(DropAllComma(txt_job_salary.Text)) & "','" & txt_job_reason.Text & "'," & _
                "'" & txt_job_description.Text & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 5 Then
        SQL = "SELECT MAX(seq_no) jmlSeq FROM m_employee_title " & _
                "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            v_seq = IIf(IsNull(rs!jmlSeq), 0, rs!jmlSeq) + 1
        Else
            v_seq = 1
        End If
        rs.Close
        
        CnG.BeginTrans
        
        SQL = "INSERT INTO m_employee_title (employee_code,seq_no,date_title,title_code,description,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & TDBGrid_Employee.Columns("employee_code").Value & "','" & v_seq & "'," & _
                "'" & Format(DTPicker_title.Value, "yyyy-MM-dd") & "','" & TDBCombo_title_emp.Text & "'," & _
                "'" & txt_description_title.Text & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 6 Then
        SQL = "SELECT MAX(seq_no) jmlSeq FROM m_employee_grade " & _
                "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            v_seq = IIf(IsNull(rs!jmlSeq), 0, rs!jmlSeq) + 1
        Else
            v_seq = 1
        End If
        rs.Close
        
        CnG.BeginTrans
        
        SQL = "INSERT INTO m_employee_grade (employee_code,seq_no,date_grade,grade_code,description,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & TDBGrid_Employee.Columns("employee_code").Value & "','" & v_seq & "'," & _
                "'" & Format(DTPicker_grade.Value, "yyyy-MM-dd") & "','" & TDBCombo_grade1.Text & "'," & _
                "'" & txt_grade_description.Text & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 7 Then
        SQL = "SELECT MAX(seq_no) jmlSeq FROM m_employee_training " & _
                "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            v_seq = IIf(IsNull(rs!jmlSeq), 0, rs!jmlSeq) + 1
        Else
            v_seq = 1
        End If
        rs.Close
        
        CnG.BeginTrans
        
        SQL = "INSERT INTO m_employee_training (employee_code,seq_no,start_date,end_date,material," & _
                "organizer,place,value,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & TDBGrid_Employee.Columns("employee_code").Value & "','" & v_seq & "'," & _
                "'" & Format(DTPicker_training_start.Value, "yyyy-MM-dd") & "','" & Format(DTPicker_training_end.Value, "yyyy-MM-dd") & "'," & _
                "'" & txt_training_subject.Text & "','" & txt_training_organize.Text & "','" & txt_training_place.Text & "'," & _
                "'" & txt_training_value.Text & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 8 Then
        SQL = "SELECT MAX(seq_no) jmlSeq FROM m_employee_contract " & _
                "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            v_seq = IIf(IsNull(rs!jmlSeq), 0, rs!jmlSeq) + 1
        Else
            v_seq = 1
        End If
        rs.Close
        
        CnG.BeginTrans
        
        SQL = "INSERT INTO m_employee_contract (employee_code,seq_no,start_date,end_date," & _
                "no_contract,description,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & TDBGrid_Employee.Columns("employee_code").Value & "','" & v_seq & "'," & _
                "'" & Format(DTPicker_contract_start.Value, "yyyy-MM-dd") & "','" & Format(DTPicker_contract_end.Value, "yyyy-MM-dd") & "'," & _
                "'" & txt_contract_no.Text & "','" & txt_contract_description & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    End If
    
    CnG.CommitTrans
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

'Private Function get_level(ByVal str_title_code As String) As Integer
'Dim rs1 As New ADODB.Recordset
'
'
'rs1.Open "select * from m_title where title_code='" & str_title_code & "'", CnG, adOpenStatic, adLockReadOnly
'If rs1.RecordCount > 0 Then
'    get_level = rs1.Fields("level").Value
'Else
'    get_level = 0
'End If
'End Function

Private Sub edit_old_data()
Dim rscari As New ADODB.Recordset

On Error GoTo Err
    CnG.BeginTrans
    If SSTab1.Tab = 0 Then
        SQL = "UPDATE m_employee_grade SET date_grade = '" & Format(DTPicker_start_working.Value, "yyyy-MM-dd") & "',grade_code = '" & TDBCombo_grade.Text & "'," & _
                "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsEmployee.Fields("employee_code").Value & "' " & _
                "AND seq_no = 1"
        CnG.Execute SQL
        
        SQL = "UPDATE m_employee SET nik = '" & txt_nik.Text & "',employee_name = '" & Replace(txt_employee_name.Text, "'", "''") & "',employee_nick_name = '" & Replace(txt_employee_nick_name, "'", "''") & "'," & _
                "company_code = '" & TDBCombo_company.Text & "',division_code = '" & TDBCombo_division.Text & "'," & _
                "title_code = '" & TDBCombo_title.Text & "',level_code = '" & TDBCombo_level.Text & "',grade_code = '" & TDBCombo_grade.Text & "',status_code = '" & TDBCombo_status.Text & "'," & _
                "place_birth = '" & txt_place_of_birth.Text & "',date_birth = '" & Format(DTPicker_birth.Value, "yyyy-MM-dd") & "',sex = '" & cbo_sex.ListIndex & "',religion = '" & cbo_religion.ListIndex & "'," & _
                "marital_status = '" & cbo_marital_status.ListIndex & "',no_of_children = '" & txt_number_of_children.Text & "',emp_address = '" & Replace(txtAlamat.Text, "'", "''") & "',npwp = '" & txt_npwp.Text & "'," & _
                "npwp_method = '" & cbo_tax_method.ListIndex & "',npwp_registered_date = '" & Format(DTPicker_tglnpwp.Value, "yyyy-MM-dd") & "',npwp_address = '" & Replace(txt_address.Text, "'", "''") & "',no_jamsostek = '" & txt_jamsostek.Text & "'," & _
                "jstk_registered_date = '" & Format(DTPicker_jstk.Value, "yyyy-MM-dd") & "',bank_code = '" & TDBCombo_bank.Text & "',bank_account = '" & txt_bank_account.Text & "',bank_acc_name = '" & Replace(txt_account_name.Text, "'", "''") & "'," & _
                "start_working = '" & Format(DTPicker_start_working.Value, "yyyy-MM-dd") & "',appointment_date = '" & Format(DTPicker_appointment.Value, "yyyy-MM-dd") & "',phone_number = '" & txt_phone_number.Text & "'," & _
                "email = '" & txt_email.Text & "',nationality = '" & cbo_nationality.ListIndex & "',identity_number = '" & txtNoKTP.Text & "',description = '" & Replace(txt_description.Text, "'", "''") & "',flag_active = '" & IIf(opt_active, 1, 0) & "'," & _
                "end_working = '" & IIf(opt_active, "0000-00-00", Format(DTPicker_end_working.Value, "yyyy-MM-dd")) & "',reason = '" & IIf(opt_active, "", txt_end_working_reason.Text) & "',picture = '" & img & "',edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsEmployee.Fields("employee_code").Value & "'"
        CnG.Execute SQL
        
        If fileExists(src) Then
            SQL = "SELECT employee_code, picture FROM m_employee WHERE employee_code = '" & rsEmployee.Fields("employee_code").Value & "'"
            If Not addImageToDB(SQL, src, "picture") Then MsgBox "Gambar Gagal Disimpan...", vbExclamation, headerMSG
        End If
    ElseIf SSTab1.Tab = 1 Then
        SQL = "UPDATE m_employee_fams SET name = '" & Replace(txt_family_name.Text, "'", "''") & "',relationship = '" & cbo_fams_rel.ListIndex & "'," & _
                "nm_rel = '" & cbo_fams_rel.Text & "',date_birth = '" & Format(DTPicker_fams_birth.Value, "yyyy-MM-dd") & "'," & _
                "sex = '" & cbo_fams_sex.ListIndex & "',education = '" & txt_fams_edu.Text & "',employment = '" & txt_fams_employment.Text & "'," & _
                "chk_address = '" & chk_fams_address.Value & "',address = '" & txt_fams_address.Text & "',edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsFams.Fields("employee_code").Value & "' " & _
                "AND seq_no = '" & rsFams.Fields("seq_no").Value & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "UPDATE m_employee_edu SET start_year = '" & Format(DTPicker_edu_start.Value, "yyyy-MM-dd") & "',end_year = '" & Format(DTPicker_edu_start.Value, "yyyy-MM-dd") & "'," & _
                "jenjang = '" & cbo_edu_level.ListIndex & "',nm_jenjang = '" & cbo_edu_level.Text & "',jurusan = '" & txt_edu_majors.Text & "',school = '" & txt_edu_school.Text & "'," & _
                "city = '" & txt_edu_city.Text & "',country = '" & txt_edu_country.Text & "',edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsEdu.Fields("employee_code").Value & "' " & _
                "AND seq_no = '" & rsEdu.Fields("seq_no").Value & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 3 Then
        SQL = "UPDATE m_employee_skill SET skill = '" & Replace(txt_skill_name.Text, "'", "''") & "',level = '" & cbo_edu_level.ListIndex & "'," & _
                "nm_level = '" & cbo_skill_level.Text & "',description = '" & txt_skill_description.Text & "'," & _
                "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsSkill.Fields("employee_code").Value & "' " & _
                "AND seq_no = '" & rsSkill.Fields("seq_no").Value & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 4 Then
        SQL = "UPDATE m_employee_exp SET start_working = '" & Format(DTPicker_job_start.Value, "yyyy-MM-dd") & "',end_working = '" & Format(DTPicker_job_end.Value, "yyyy-MM-dd") & "'," & _
                "company_name = '" & txt_job_company.Text & "',usaha = '" & txt_job_line.Text & "',department_name = '" & txt_job_dept.Text & "',last_title = '" & txt_job_title.Text & "'," & _
                "last_salary = '" & Val(DropAllComma(txt_job_salary.Text)) & "',reason_stop_working = '" & txt_job_reason.Text & "',description = '" & txt_job_description.Text & "',edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsJob.Fields("employee_code").Value & "' " & _
                "AND seq_no = '" & rsJob.Fields("seq_no").Value & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 5 Then
        SQL = "UPDATE m_employee_title SET date_title = '" & Format(DTPicker_title.Value, "yyyy-MM-dd") & "',title_code = '" & TDBCombo_title_emp.Text & "'," & _
                "description = '" & txt_description_title.Text & "',edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsTitleEmp.Fields("employee_code").Value & "' " & _
                "AND seq_no = '" & rsTitleEmp.Fields("seq_no").Value & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 6 Then
        SQL = "UPDATE m_employee_grade SET date_grade = '" & Format(DTPicker_grade.Value, "yyyy-MM-dd") & "',grade_code = '" & TDBCombo_grade1.Text & "'," & _
                "description = '" & txt_grade_description.Text & "',edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsGrade1.Fields("employee_code").Value & "' " & _
                "AND seq_no = '" & rsGrade1.Fields("seq_no").Value & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 7 Then
        SQL = "UPDATE m_employee_training SET start_date = '" & Format(DTPicker_training_start.Value, "yyyy-MM-dd") & "',end_date = '" & Format(DTPicker_training_end.Value, "yyyy-MM-dd") & "'," & _
                "material = '" & txt_training_subject.Text & "',organizer = '" & txt_training_organize.Text & "',place = '" & txt_training_place.Text & "',value = '" & txt_training_value.Text & "'," & _
                "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsTraining.Fields("employee_code").Value & "' " & _
                "AND seq_no = '" & rsTraining.Fields("seq_no").Value & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 8 Then
        SQL = "UPDATE m_employee_contract SET start_date = '" & Format(DTPicker_contract_start.Value, "yyyy-MM-dd") & "',end_date = '" & Format(DTPicker_contract_end.Value, "yyyy-MM-dd") & "'," & _
                "no_contract = '" & txt_contract_no.Text & "',description = '" & txt_contract_description.Text & "',edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsContract.Fields("employee_code").Value & "' " & _
                "AND seq_no = '" & rsContract.Fields("seq_no").Value & "'"
        CnG.Execute SQL
    End If
    
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
        If check_validate_exist_new Then
            Call check_invalid: Exit Sub
        End If
        Call insert_new_data
        
        If SSTab1.Tab = 0 Then
            clsFunc.InsertLog ("Insert Master Employee : " & txt_nik.Text)
        End If
    ElseIf int_mode = 2 Then
        If Not check_validate_new Then Exit Sub
    '    If check_validate_exist_edit Then
    '        Call check_invalid: Exit Sub
    '    End If
        
        If SSTab1.Tab = 0 Then
            If v_company <> TDBCombo_company.Text Then
                If Not check_validate_new Then Exit Sub
                    If check_validate_exist_new Then
                        Call check_invalid: Exit Sub
                    End If
                Call insert_new_data
            Else
                Call edit_old_data
            End If
            clsFunc.InsertLog ("Edit Master Employee : " & txt_nik.Text)
        Else
            Call edit_old_data
        End If
    End If
    
    If SSTab1.Tab = 0 Then
        Call load_data_employee
        fra_status_emp.Visible = True
    ElseIf SSTab1.Tab = 1 Then
        Call load_data_employee_family
    ElseIf SSTab1.Tab = 2 Then
        Call load_data_employee_education
    ElseIf SSTab1.Tab = 3 Then
        Call load_data_employee_skill
    ElseIf SSTab1.Tab = 4 Then
        Call load_data_employee_job
    ElseIf SSTab1.Tab = 5 Then
        Call load_data_employee_title
    ElseIf SSTab1.Tab = 6 Then
        Call load_data_employee_grade
    ElseIf SSTab1.Tab = 7 Then
        Call load_data_employee_training
    ElseIf SSTab1.Tab = 8 Then
        Call load_data_employee_contract
    End If
    
    int_mode = 0
    Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
    If SSTab1.Tab = 0 Then
        cmdNew(0).Enabled = a And blnUser_Add
        cmdSave(0).Enabled = b
        cmdEdit(0).Enabled = c And blnUser_Edit
        cmdDelete(0).Enabled = d And blnUser_Delete
        cmdCancel(0).Enabled = e
    ElseIf SSTab1.Tab = 1 Then
        cmdNew(1).Enabled = a And blnUser_Add
        cmdSave(1).Enabled = b
        cmdEdit(1).Enabled = c And blnUser_Edit
        cmdDelete(1).Enabled = d And blnUser_Delete
        cmdCancel(1).Enabled = e
    ElseIf SSTab1.Tab = 2 Then
        cmdNew(2).Enabled = a And blnUser_Add
        cmdSave(2).Enabled = b
        cmdEdit(2).Enabled = c And blnUser_Edit
        cmdDelete(2).Enabled = d And blnUser_Delete
        cmdCancel(2).Enabled = e
    ElseIf SSTab1.Tab = 3 Then
        cmdNew(3).Enabled = a And blnUser_Add
        cmdSave(3).Enabled = b
        cmdEdit(3).Enabled = c And blnUser_Edit
        cmdDelete(3).Enabled = d And blnUser_Delete
        cmdCancel(3).Enabled = e
    ElseIf SSTab1.Tab = 4 Then
        cmdNew(4).Enabled = a And blnUser_Add
        cmdSave(4).Enabled = b
        cmdEdit(4).Enabled = c And blnUser_Edit
        cmdDelete(4).Enabled = d And blnUser_Delete
        cmdCancel(4).Enabled = e
    ElseIf SSTab1.Tab = 5 Then
        cmdNew(8).Enabled = a And blnUser_Add
        cmdSave(8).Enabled = b
        cmdEdit(8).Enabled = c And blnUser_Edit
        cmdDelete(8).Enabled = d And blnUser_Delete
        cmdCancel(8).Enabled = e
    ElseIf SSTab1.Tab = 6 Then
        cmdNew(5).Enabled = a And blnUser_Add
        cmdSave(5).Enabled = b
        cmdEdit(5).Enabled = c And blnUser_Edit
        cmdDelete(5).Enabled = d And blnUser_Delete
        cmdCancel(5).Enabled = e
    ElseIf SSTab1.Tab = 7 Then
        cmdNew(6).Enabled = a And blnUser_Add
        cmdSave(6).Enabled = b
        cmdEdit(6).Enabled = c And blnUser_Edit
        cmdDelete(6).Enabled = d And blnUser_Delete
        cmdCancel(6).Enabled = e
    ElseIf SSTab1.Tab = 8 Then
        cmdNew(7).Enabled = a And blnUser_Add
        cmdSave(7).Enabled = b
        cmdEdit(7).Enabled = c And blnUser_Edit
        cmdDelete(7).Enabled = d And blnUser_Delete
        cmdCancel(7).Enabled = e
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
        ElseIf TypeOf Ctr Is Image Then
            If Not LCase(Ctr.name) = "image1" Then Ctr.Picture = Nothing
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
        cbo_sex.ListIndex = 0
        cbo_religion.ListIndex = 0
        cbo_marital_status.ListIndex = 0
        cbo_nationality.ListIndex = 0
        cbo_tax_method.ListIndex = 0
        
        TDBCombo_title.Enabled = True
        TDBCombo_grade.Enabled = True
        
        DTPicker_birth.Value = Now
        DTPicker_appointment.Value = Now
        DTPicker_start_working.Value = Now
        DTPicker_tglnpwp.Value = Now
        DTPicker_jstk.Value = Now
        DTPicker_end_working.Value = Now
        
        txt_number_of_children.Text = 0
        
        '--------------------------- Employee Picture Default -----------------------
        txt_pict_location.Text = App.Path & "\anonymous.jpg"
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
        
        Call set_age_data
        Call set_working_age_data
    ElseIf SSTab1.Tab = 1 Then
        cbo_fams_rel.ListIndex = 0
        cbo_fams_sex.ListIndex = 0
        
        chk_fams_address.Value = 0
        
        DTPicker_fams_birth.Value = Now
    ElseIf SSTab1.Tab = 2 Then
        cbo_edu_level.ListIndex = 0
        
        DTPicker_edu_start.Value = Now
        DTPicker_edu_end.Value = Now
    ElseIf SSTab1.Tab = 3 Then
        cbo_skill_level.ListIndex = 0
    ElseIf SSTab1.Tab = 4 Then
        DTPicker_job_start.Value = Now
        DTPicker_job_end.Value = Now
    ElseIf SSTab1.Tab = 5 Then
        DTPicker_title.Value = Now
    ElseIf SSTab1.Tab = 6 Then
        DTPicker_grade.Value = Now
    ElseIf SSTab1.Tab = 7 Then
        DTPicker_training_start.Value = Now
        DTPicker_training_end.Value = Now
    ElseIf SSTab1.Tab = 8 Then
        DTPicker_contract_start.Value = Now
        DTPicker_contract_end.Value = Now
    End If
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        Call set_new_data
        
        If SSTab1.Tab = 0 Then
            fra_entry_employee.Visible = True
            txt_nik.Enabled = True
            TDBGrid_Employee.Enabled = False
            
            If txt_nik.Enabled = True Then
                txt_nik.SetFocus
            End If
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_family.Visible = True
            txt_family_name.Enabled = True
            TDBGrid_Family.Enabled = False
            
            If txt_family_name.Enabled = True Then
                txt_family_name.SetFocus
            End If
        ElseIf SSTab1.Tab = 2 Then
            fra_entry_education.Visible = True
            txt_edu_majors.Enabled = True
            TDBGrid_Education.Enabled = False
            
            If txt_edu_majors.Enabled = True Then
                txt_edu_majors.SetFocus
            End If
        ElseIf SSTab1.Tab = 3 Then
            fra_entry_skill.Visible = True
            txt_skill_name.Enabled = True
            TDBGrid_Skill.Enabled = False
            
            If txt_skill_name.Enabled = True Then
                txt_skill_name.SetFocus
            End If
        ElseIf SSTab1.Tab = 4 Then
            fra_entry_job.Visible = True
            txt_job_company.Enabled = True
            TDBGrid_Job.Enabled = False
            
            If txt_job_company.Enabled = True Then
                txt_job_company.SetFocus
            End If
        ElseIf SSTab1.Tab = 5 Then
            fra_entry_title.Visible = True
            TDBCombo_title_emp.Enabled = True
            TDBGrid_Title.Enabled = False
            
            If TDBCombo_title_emp.Enabled = True Then
                TDBCombo_title_emp.SetFocus
            End If
        ElseIf SSTab1.Tab = 6 Then
            fra_entry_grade.Visible = True
            TDBCombo_grade1.Enabled = True
            TDBGrid_Grade.Enabled = False
            
            If TDBCombo_grade1.Enabled = True Then
                TDBCombo_grade1.SetFocus
            End If
        ElseIf SSTab1.Tab = 7 Then
            fra_entry_training.Visible = True
            txt_training_subject.Enabled = True
            TDBGrid_Training.Enabled = False
            
            If txt_training_subject.Enabled = True Then
                txt_training_subject.SetFocus
            End If
        ElseIf SSTab1.Tab = 8 Then
            fra_entry_contract.Visible = True
            txt_contract_no.Enabled = True
            TDBGrid_Contract.Enabled = False
            
            If txt_contract_no.Enabled = True Then
                txt_contract_no.SetFocus
            End If
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        
        If SSTab1.Tab = 0 Then
            fra_entry_employee.Visible = False
            TDBGrid_Employee.Enabled = True
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_family.Visible = False
            TDBGrid_Family.Enabled = True
        ElseIf SSTab1.Tab = 2 Then
            fra_entry_education.Visible = False
            TDBGrid_Education.Enabled = True
        ElseIf SSTab1.Tab = 3 Then
            fra_entry_skill.Visible = False
            TDBGrid_Skill.Enabled = True
        ElseIf SSTab1.Tab = 4 Then
            fra_entry_job.Visible = False
            TDBGrid_Job.Enabled = True
        ElseIf SSTab1.Tab = 5 Then
            fra_entry_title.Visible = False
            TDBGrid_Title.Enabled = True
        ElseIf SSTab1.Tab = 6 Then
            fra_entry_grade.Visible = False
            TDBGrid_Grade.Enabled = True
        ElseIf SSTab1.Tab = 7 Then
            fra_entry_training.Visible = False
            TDBGrid_Training.Enabled = True
        ElseIf SSTab1.Tab = 8 Then
            fra_entry_contract.Visible = False
            TDBGrid_Contract.Enabled = True
        End If
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        If SSTab1.Tab = 0 Then
'            txt_nik.Enabled = False
            fra_entry_employee.Visible = True
            TDBGrid_Employee.Enabled = False
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_family.Visible = True
            TDBGrid_Family.Enabled = False
        ElseIf SSTab1.Tab = 2 Then
            fra_entry_education.Visible = True
            TDBGrid_Education.Enabled = False
        ElseIf SSTab1.Tab = 3 Then
            fra_entry_skill.Visible = True
            TDBGrid_Skill.Enabled = False
        ElseIf SSTab1.Tab = 4 Then
            fra_entry_job.Visible = True
            TDBGrid_Job.Enabled = False
        ElseIf SSTab1.Tab = 5 Then
            fra_entry_title.Visible = True
            TDBGrid_Title.Enabled = False
        ElseIf SSTab1.Tab = 6 Then
            fra_entry_grade.Visible = True
            TDBGrid_Grade.Enabled = False
        ElseIf SSTab1.Tab = 7 Then
            fra_entry_training.Visible = True
            TDBGrid_Training.Enabled = False
        ElseIf SSTab1.Tab = 8 Then
            fra_entry_contract.Visible = True
            TDBGrid_Contract.Enabled = False
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

Private Sub set_age_data()
    txt_age = Trim(str(year(Now) - year(DTPicker_birth.Value)))
End Sub

Private Sub set_working_age_data()
    txt_working_time = Trim(str(year(Now) - year(DTPicker_start_working.Value)))
End Sub

Private Sub chk_fams_address_Click()
    If chk_fams_address.Value = 0 Then
        txt_fams_address.Text = ""
        txt_fams_address.Enabled = True
        txt_fams_address.SetFocus
    Else
        txt_fams_address.Text = TDBGrid_Employee.Columns("emp_address").Value
        txt_fams_address.Enabled = False
    End If
End Sub

Private Sub cmdImport_Click()
    frm_import_employee.Show
End Sub

Private Sub cmdPrint_Click()
Dim str_param_periode As String
Dim a As New frm_rpt
Dim str_file As String
Dim lama_kerja As String
Dim lama_usia As String

    lama_kerja = Trim(str(year(Now) - year(TDBGrid_Employee.Columns("start_working").Value)))
    lama_usia = Trim(str(year(Now) - year(TDBGrid_Employee.Columns("date_birth").Value)))
    
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
    str_param_periode = ""
    
    SQL = "SELECT a.company_code, b.company_name, a.division_code, c.division_name, " & _
            "a.employee_code, a.nik, a.employee_name, a.title_code, d.title_name, " & _
            "a.start_working,'" & lama_kerja & "',e.status_name,place_birth, date_birth, " & _
            "'" & lama_usia & "',sex,religion, marital_status, no_of_children, emp_address, " & _
            "a.phone_number, email, identity_number, nationality, a.description, a.npwp, npwp_registered_date, npwp_address, " & _
            "no_jamsostek, jstk_registered_date, bank_code, bank_account, bank_acc_name, " & _
            "flag_active, reason, '" & IIf(LOGIN_LEVEL = 100, LOGIN_FULLNAME, EMPLOYEE_NAME) & "' " & _
          "FROM m_employee a JOIN m_company b ON a.company_code = b.company_code " & _
            "JOIN m_division c ON a.division_code = c.division_code AND a.company_code = c.company_code " & _
            "JOIN m_title d ON a.title_code = d.title_code " & _
            "JOIN m_emp_status e ON a.status_code = e.status_code " & _
          "WHERE employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "'"
    str_file = "\report\rpt_emp_detail.rpt"
    
    Call a.Show
    a.Caption = "DETAIL KARYAWAN"
     
    Call a.rpt_view(SQL, str_file, str_param_periode)

End Sub

Private Sub DTPicker_birth_Change()
    Call set_age_data
End Sub

Private Sub DTPicker_start_working_Change()
    Call set_working_age_data
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
        
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False
    SSTab1.TabEnabled(6) = False
    SSTab1.TabEnabled(7) = False
    SSTab1.TabEnabled(8) = False
    
    If SSTab1.Tab = 0 Then
        Call load_data_company
        Call load_data_title
        Call load_data_level
        Call load_data_grade
        Call load_data_status
        Call load_data_bank
        
        fra_entry_employee.Visible = True
    End If
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_employee = Nothing
End Sub

Private Sub opt_active_Click()
Dim v_flag_active As Integer

    SQL = "SELECT flag_active FROM m_employee WHERE employee_code = '" & txt_employee_code & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
    If Not rs.EOF Then
        v_flag_active = rs!flag_active
    End If
        
    If rs.RecordCount > 0 Then
        If v_flag_active = 1 Then
            fra_not_active.Visible = False
            DTPicker_end_working.Value = Now
            txt_end_working_reason.Text = ""
        End If
    End If
    rs.Close
End Sub

Private Sub opt_not_active_Click()
    If opt_not_active Then
        fra_not_active.Visible = True
        DTPicker_end_working.Value = Now
        lbl_reason_end_working.Caption = "End Working"
        txt_end_working_reason.Text = ""
    End If
End Sub

Private Sub optActive_Click()
    Call load_data_employee
    Call load_count_employee
End Sub

Private Sub optNotActive_Click()
    Call load_data_employee
    Call load_count_employee
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    SQL = "SELECT b.flag_period FROM m_employee a JOIN m_emp_status b ON a.status_code = b.status_code " & _
          "WHERE a.employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        v_flag_periode = rs!flag_period
    End If
    rs.Close
    
    If v_flag_periode = 0 Then
        TDBGrid_Contract.Enabled = False
        Frame5(3).Enabled = False
    Else
        TDBGrid_Contract.Enabled = True
        Frame5(3).Enabled = True
    End If
    
    If SSTab1.Tab <> 0 Then
        If Not (TDBGrid_Employee.ApproxCount > 0 And TDBGrid_Employee.Bookmark > 0) Then
            lbl_fam.Caption = ""
            lbl_edu.Caption = ""
            lbl_skill.Caption = ""
            lbl_exp.Caption = ""
            lbl_jbtn.Caption = ""
            lbl_grade.Caption = ""
            lbl_training.Caption = ""
            lbl_kontrak.Caption = ""
        Else
            lbl_fam.Caption = "NAMA KARYAWAN : " & TDBGrid_Employee.Columns("employee_name").Value
            lbl_edu.Caption = "NAMA KARYAWAN : " & TDBGrid_Employee.Columns("employee_name").Value
            lbl_skill.Caption = "NAMA KARYAWAN : " & TDBGrid_Employee.Columns("employee_name").Value
            lbl_exp.Caption = "NAMA KARYAWAN : " & TDBGrid_Employee.Columns("employee_name").Value
            lbl_jbtn.Caption = "NAMA KARYAWAN : " & TDBGrid_Employee.Columns("employee_name").Value
            lbl_grade.Caption = "NAMA KARYAWAN : " & TDBGrid_Employee.Columns("employee_name").Value
            lbl_training.Caption = "NAMA KARYAWAN : " & TDBGrid_Employee.Columns("employee_name").Value
            lbl_kontrak.Caption = "NAMA KARYAWAN : " & TDBGrid_Employee.Columns("employee_name").Value
        End If
    End If
    
    int_mode = 0
    Call load_mode
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
        
        lbl_employee.Visible = True
        optActive.Value = True
        Call load_data_employee
        Call load_data_division
        Call load_count_employee
    End If
End Sub

'Private Sub TDBCombo_department_itemChange()
'    If TDBCombo_department.ApproxCount > 0 Then
'        TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
'        txt_department_name.Text = TDBCombo_department.Columns("department_name").Value
'        If int_mode = 1 Or int_mode = 2 Then _
'            Call load_data_division(TDBCombo_department.Columns("department_code").Value)
'    End If
'End Sub

Private Sub tdbcombo_division_itemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
        txt_division_name.Text = TDBCombo_division.Columns("division_name").Value
    End If
End Sub

Private Sub TDBCombo_title_ItemChange()
    If TDBCombo_title.ApproxCount > 0 Then
        TDBCombo_title.Text = TDBCombo_title.Columns("title_code").Value
        txt_title_name.Text = TDBCombo_title.Columns("title_name").Value
    End If
End Sub

Private Sub TDBCombo_level_ItemChange()
    If TDBCombo_level.ApproxCount > 0 Then
        TDBCombo_level.Text = TDBCombo_level.Columns("level_code").Value
        txt_level_name.Text = TDBCombo_level.Columns("level_name").Value
    End If
End Sub

Private Sub TDBCombo_grade_ItemChange()
    If TDBCombo_grade.ApproxCount > 0 Then
        TDBCombo_grade.Text = TDBCombo_grade.Columns("grade_code").Value
        txt_grade_name.Text = TDBCombo_grade.Columns("grade_name").Value
    End If
End Sub

Private Sub TDBCombo_title_emp_ItemChange()
    If TDBCombo_title_emp.ApproxCount > 0 Then
        TDBCombo_title_emp.Text = TDBCombo_title_emp.Columns("title_code").Value
        txt_title.Text = TDBCombo_title_emp.Columns("title_name").Value
    End If
End Sub

Private Sub TDBCombo_grade1_ItemChange()
    If TDBCombo_grade1.ApproxCount > 0 Then
        TDBCombo_grade1.Text = TDBCombo_grade1.Columns("grade_code").Value
        txt_grade_name1.Text = TDBCombo_grade1.Columns("grade_name").Value
    End If
End Sub

Private Sub TDBCombo_status_ItemChange()
    If TDBCombo_status.ApproxCount > 0 Then
        TDBCombo_status.Text = TDBCombo_status.Columns("status_code").Value
        txt_status_name.Text = TDBCombo_status.Columns("status_name").Value
    End If
End Sub

Private Sub TDBCombo_bank_ItemChange()
    If TDBCombo_bank.ApproxCount > 0 Then
        TDBCombo_bank.Text = TDBCombo_bank.Columns("bank_code").Value
        txt_bank_name.Text = TDBCombo_bank.Columns("bank_name").Value
        
        txt_account_name.Text = txt_employee_name.Text
    End If
End Sub

Public Sub load_data_employee()
    If rsEmployee.State Then rsEmployee.Close
    If LOGIN_LEVEL = 100 Then
        If optActive.Value = True Then
            SQL = "SELECT a.*,c.division_name,d.title_name " & _
                  "FROM m_employee a JOIN m_division c ON a.division_code = c.division_code " & _
                    "LEFT JOIN m_title d ON a.title_code = d.title_code " & _
                  "WHERE a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                    "AND flag_active <> 0 order by employee_name"
        Else
            SQL = "SELECT a.*,c.division_name,d.title_name " & _
                  "FROM m_employee a JOIN m_division c ON a.division_code = c.division_code " & _
                    "LEFT JOIN m_title d ON a.title_code = d.title_code " & _
                  "WHERE a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                    "AND flag_active = 0 order by employee_name"
        End If
    Else
        If optActive.Value = True Then
            SQL = "SELECT a.*,c.division_name,d.title_name " & _
                  "FROM m_employee a JOIN m_division c ON a.division_code = c.division_code " & _
                    "LEFT JOIN m_title d ON a.title_code = d.title_code " & _
                  "WHERE a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                    "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND flag_active <> 0 order by employee_name"
        Else
            SQL = "SELECT a.*,c.division_name,d.title_name " & _
                  "FROM m_employee a JOIN m_division c ON a.division_code = c.division_code " & _
                    "LEFT JOIN m_title d ON a.title_code = d.title_code " & _
                  "WHERE a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                    "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND flag_active = 0 order by employee_name"
        End If
    End If

    rsEmployee.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    fra_status_emp.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)

    TDBGrid_Employee.DataSource = rsEmployee
    
    If IsNull(TDBGrid_Employee.Columns("employee_code").Value) Then
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        SSTab1.TabEnabled(5) = False
        SSTab1.TabEnabled(6) = False
        SSTab1.TabEnabled(7) = False
        SSTab1.TabEnabled(8) = False
    Else
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = True
        SSTab1.TabEnabled(4) = True
        SSTab1.TabEnabled(5) = True
        SSTab1.TabEnabled(6) = True
        SSTab1.TabEnabled(7) = True
        SSTab1.TabEnabled(8) = True
    End If
End Sub

Public Sub load_data_company()
    TDBCombo_company.Text = "": txt_company_name = ""
    
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

'Private Sub load_data_department()
'    TDBCombo_department.Text = "": txt_department_name = ""
'
'    If rsDepartment.State Then rsDepartment.Close
'    SQL = "select * from m_department where company_code='" _
'            & TDBCombo_company.Columns("company_code").Value & "' order by department_code"
'    rsDepartment.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'    TDBCombo_department.RowSource = rsDepartment
'End Sub

Private Sub load_data_division()
    'TDBCombo_division.Text = "": txt_division_name = ""
    
    If rsDivision.State Then rsDivision.Close
    SQL = "select * from m_division where company_code='" _
            & TDBCombo_company.Columns("company_code").Value & "' order by division_code"
    rsDivision.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_division.RowSource = rsDivision
End Sub

Private Sub load_data_title()
    If rsTitle.State Then rsTitle.Close
    SQL = "select * from m_title order by title_code"
    rsTitle.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_title.RowSource = rsTitle
End Sub

Private Sub load_data_level()
    If rsLevel.State Then rsLevel.Close
    SQL = "select * from m_level order by level_code"
    rsLevel.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_level.RowSource = rsLevel
End Sub

Private Sub load_data_grade()
    If rsGrade.State Then rsGrade.Close
    SQL = "select * from m_grade order by grade_code"
    rsGrade.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_grade.RowSource = rsGrade
End Sub

Private Sub load_data_title_emp()
    If rsTDBTitle.State Then rsTDBTitle.Close
    SQL = "select * from m_title order by title_code"
    rsTDBTitle.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_title_emp.RowSource = rsTDBTitle
End Sub

Private Sub load_data_grade1()
    If rsTDBGrade.State Then rsTDBGrade.Close
    SQL = "select * from m_grade order by grade_code"
    rsTDBGrade.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_grade1.RowSource = rsTDBGrade
End Sub

Private Sub load_data_status()
    If rsStatus.State Then rsStatus.Close
    SQL = "select * from m_emp_status order by status_code"
    rsStatus.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_status.RowSource = rsStatus
End Sub

Private Sub load_data_bank()
    If rsBank.State Then rsBank.Close
    SQL = "select * from m_bank order by bank_code"
    rsBank.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_bank.RowSource = rsBank
End Sub

Private Sub load_data_employee_family()
    If rsFams.State Then rsFams.Close
    SQL = "select *,case when sex = 0 then 'Male' else 'Female' end AS jenkel " & _
          "from m_employee_fams " & _
          "where employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
          "order by seq_no"
    rsFams.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Family.DataSource = rsFams
End Sub

Private Sub load_data_employee_education()
    If rsEdu.State Then rsEdu.Close
    SQL = "select * from m_employee_edu " & _
          "where employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
          "order by seq_no"
    rsEdu.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Education.DataSource = rsEdu
End Sub

Private Sub load_data_employee_skill()
    If rsSkill.State Then rsSkill.Close
    SQL = "select * from m_employee_skill " & _
          "where employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
          "order by seq_no"
    rsSkill.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Skill.DataSource = rsSkill
End Sub

Private Sub load_data_employee_job()
    If rsJob.State Then rsJob.Close
    SQL = "select * from m_employee_exp " & _
          "where employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
          "order by seq_no"
    rsJob.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Job.DataSource = rsJob
End Sub

Private Sub load_data_employee_training()
    If rsTraining.State Then rsTraining.Close
    SQL = "select * from m_employee_training " & _
          "where employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
          "order by seq_no"
    rsTraining.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Training.DataSource = rsTraining
End Sub

Private Sub load_data_employee_contract()
    If rsContract.State Then rsContract.Close
    SQL = "select * from m_employee_contract " & _
          "where employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
          "order by seq_no"
    rsContract.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Contract.DataSource = rsContract
End Sub

Private Sub load_data_employee_title()
    If rsTitleEmp.State Then rsTitleEmp.Close
    SQL = "select a.*,b.title_name from m_employee_title a join m_title b on a.title_code = b.title_code " & _
          "where a.employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
          "order by a.seq_no"
    rsTitleEmp.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Title.DataSource = rsTitleEmp
End Sub

Private Sub load_data_employee_grade()
    If rsGrade1.State Then rsGrade1.Close
    SQL = "select a.*,b.grade_name from m_employee_grade a join m_grade b on a.grade_code = b.grade_code " & _
          "where a.employee_code = '" & TDBGrid_Employee.Columns("employee_code").Value & "' " & _
          "order by a.seq_no"
    rsGrade1.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Grade.DataSource = rsGrade1
End Sub

Private Sub TDBGrid_employee_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If TDBGrid_Employee.Columns(ColIndex).Caption = "BIRTH DATE" Or _
    TDBGrid_Employee.Columns(ColIndex).Caption = "START WORKING" Or _
    TDBGrid_Employee.Columns(ColIndex).Caption = "END WORKING" Then
        Value = Format(Value, "yyyy-mm-dd")
    End If
End Sub

Private Sub TDBGrid_Employee_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call load_data_employee_family
    Call load_data_employee_education
    Call load_data_employee_skill
    Call load_data_employee_job
    Call load_data_title_emp
    Call load_data_employee_title
    Call load_data_grade1
    Call load_data_employee_grade
    Call load_data_employee_training
    Call load_data_employee_contract
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Public Sub load_count_employee()
Dim active As String

    active = IIf(optActive.Value, "<> 0", "= 0")
    
    If rs.State Then rs.Close
    If LOGIN_LEVEL = 100 Then
        SQL = "Select Count(employee_code) jml_emp From m_employee " _
                & "WHERE company_code = '" & TDBCombo_company.Text & "' " _
                & "AND " & IIf(optActive = True, "flag_active <> 0", "flag_active = 0") & ""
    Else
        SQL = "Select Count(employee_code) jml_emp From m_employee " _
                & "WHERE company_code = '" & TDBCombo_company.Text & "' " _
                & "AND " & IIf(optActive = True, "flag_active <> 0", "flag_active = 0") & " " _
                & "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND flag_active " & active & ""
    End If
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        lbl_employee.Caption = "Total Karyawan = " & rs!jml_emp
    End If
    rs.Close

End Sub

Private Sub cmd_brows_pict_Click()
Dim cls As New clsDlg
Dim i As Double
    src = cls.OpenFlDlg(Me.hwnd, "Images File(*.jpg)|*.jpg", "Open File", vbNullString, True)
    
    If src <> "" Then
        i = Round(FileLen(src) / 1024, 0)
    
        If i > 100 Then
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

'
'    If src <> "" Then
''        dst = App.Path & "\employee_pict\" & Replace(txt_nik.Text, " ", "") & "-" & Replace(txt_employee_name.Text, " ", "-") & ".jpg"
'        txt_pict_location.Text = App.Path & "\employee_pict\" & Replace(txt_nik.Text, " ", "") & "-" & Replace(txt_employee_name.Text, " ", "-") & ".jpg"
'
'        If copyFileFolder(src, txt_pict_location) = True Then
'            img.Picture = LoadPicture(txt_pict_location)
'
'            img.Width = img.Picture.Width
'            img.Height = img.Picture.Height
'            If pic.Width < img.Width Then
'                img.Width = pic.Width
'    '            img.Height = img.Height / (img.Picture.Width / img.Width)
'            End If
'
'            If pic.Height < img.Height Then
'                img.Height = pic.Height
'    '            img.Width = img.Width / (img.Picture.Height / img.Height)
'            End If
'
'            img.Left = 0
'            img.Top = 0
'        End If
'    Else
'        txt_pict_location.Text = ""
'    End If
End Sub

Private Function copyFileFolder(ByVal src$, ByVal dst) As Boolean
On Error GoTo Err

    copyFileFolder = True

    Dim SHFileOp As SHFILEOPSTRUCT
    With SHFileOp
        'Copy the file
        .wFunc = FO_COPY
        'Source the file
        .pFrom = src
        'Destination the file
        .pTo = dst
        'Allow 'move to recycle bn'
        .fFlags = FOF_ALLOWUNDO
    End With
    'perform file operation
    SHFileOperation SHFileOp
    'MsgBox "The file or folder '" + Src + "' has been copyed to destination !", vbInformation + vbOKOnly, App.Title
    Exit Function
    
Err:
copyFileFolder = False
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Function

Private Sub clear_filter()
    If SSTab1.Tab = 0 Then
        For Each Col In TDBGrid_Employee.Columns
            Col.FilterText = ""
        Next Col
        rsEmployee.Filter = adFilterNone
    ElseIf SSTab1.Tab = 1 Then
        For Each Col In TDBGrid_Family.Columns
            Col.FilterText = ""
        Next Col
        rsFams.Filter = adFilterNone
    ElseIf SSTab1.Tab = 2 Then
        For Each Col In TDBGrid_Education.Columns
            Col.FilterText = ""
        Next Col
        rsEdu.Filter = adFilterNone
    ElseIf SSTab1.Tab = 3 Then
        For Each Col In TDBGrid_Skill.Columns
            Col.FilterText = ""
        Next Col
        rsSkill.Filter = adFilterNone
    ElseIf SSTab1.Tab = 4 Then
        For Each Col In TDBGrid_Job.Columns
            Col.FilterText = ""
        Next Col
        rsJob.Filter = adFilterNone
    ElseIf SSTab1.Tab = 5 Then
        For Each Col In TDBGrid_Title.Columns
            Col.FilterText = ""
        Next Col
        rsTitleEmp.Filter = adFilterNone
    ElseIf SSTab1.Tab = 6 Then
        For Each Col In TDBGrid_Grade.Columns
            Col.FilterText = ""
        Next Col
        rsGrade1.Filter = adFilterNone
    ElseIf SSTab1.Tab = 7 Then
        For Each Col In TDBGrid_Training.Columns
            Col.FilterText = ""
        Next Col
        rsTraining.Filter = adFilterNone
    ElseIf SSTab1.Tab = 8 Then
        For Each Col In TDBGrid_Contract.Columns
            Col.FilterText = ""
        Next Col
        rsContract.Filter = adFilterNone
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
        Set Cols = TDBGrid_Employee.Columns
        i = TDBGrid_Employee.Col
        TDBGrid_Employee.HoldFields
        
        rsEmployee.Filter = getFilter()
        TDBGrid_Employee.Col = i
        TDBGrid_Employee.EditActive = True
        
        TDBGrid_Employee.SelStart = Len(TDBGrid_Employee.Columns(i).FilterText)
        If TDBGrid_Employee.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Employee.Col = i
        End If
    ElseIf SSTab1.Tab = 1 Then
        Set Cols = TDBGrid_Family.Columns
        i = TDBGrid_Family.Col
        TDBGrid_Family.HoldFields
        
        rsFams.Filter = getFilter()
        TDBGrid_Family.Col = i
        TDBGrid_Family.EditActive = True
        
        TDBGrid_Family.SelStart = Len(TDBGrid_Family.Columns(i).FilterText)
        If TDBGrid_Family.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Family.Col = i
        End If
    ElseIf SSTab1.Tab = 2 Then
        Set Cols = TDBGrid_Education.Columns
        i = TDBGrid_Education.Col
        TDBGrid_Education.HoldFields
        
        rsEdu.Filter = getFilter()
        TDBGrid_Education.Col = i
        TDBGrid_Education.EditActive = True
        
        TDBGrid_Education.SelStart = Len(TDBGrid_Education.Columns(i).FilterText)
        If TDBGrid_Education.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Education.Col = i
        End If
    ElseIf SSTab1.Tab = 3 Then
        Set Cols = TDBGrid_Skill.Columns
        i = TDBGrid_Skill.Col
        TDBGrid_Skill.HoldFields
        
        rsSkill.Filter = getFilter()
        TDBGrid_Skill.Col = i
        TDBGrid_Skill.EditActive = True
        
        TDBGrid_Skill.SelStart = Len(TDBGrid_Skill.Columns(i).FilterText)
        If TDBGrid_Skill.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Skill.Col = i
        End If
    ElseIf SSTab1.Tab = 4 Then
        Set Cols = TDBGrid_Job.Columns
        i = TDBGrid_Job.Col
        TDBGrid_Job.HoldFields
        
        rsJob.Filter = getFilter()
        TDBGrid_Job.Col = i
        TDBGrid_Job.EditActive = True
        
        TDBGrid_Job.SelStart = Len(TDBGrid_Job.Columns(i).FilterText)
        If TDBGrid_Job.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Job.Col = i
        End If
    ElseIf SSTab1.Tab = 5 Then
        Set Cols = TDBGrid_Title.Columns
        i = TDBGrid_Title.Col
        TDBGrid_Title.HoldFields
        
        rsTitleEmp.Filter = getFilter()
        TDBGrid_Title.Col = i
        TDBGrid_Title.EditActive = True
        
        TDBGrid_Title.SelStart = Len(TDBGrid_Title.Columns(i).FilterText)
        If TDBGrid_Title.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Title.Col = i
        End If
    ElseIf SSTab1.Tab = 6 Then
        Set Cols = TDBGrid_Grade.Columns
        i = TDBGrid_Grade.Col
        TDBGrid_Grade.HoldFields
        
        rsGrade1.Filter = getFilter()
        TDBGrid_Grade.Col = i
        TDBGrid_Grade.EditActive = True
        
        TDBGrid_Grade.SelStart = Len(TDBGrid_Grade.Columns(i).FilterText)
        If TDBGrid_Grade.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Grade.Col = i
        End If
    ElseIf SSTab1.Tab = 7 Then
        Set Cols = TDBGrid_Training.Columns
        i = TDBGrid_Training.Col
        TDBGrid_Training.HoldFields
        
        rsTraining.Filter = getFilter()
        TDBGrid_Training.Col = i
        TDBGrid_Training.EditActive = True
        
        TDBGrid_Training.SelStart = Len(TDBGrid_Training.Columns(i).FilterText)
        If TDBGrid_Training.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Training.Col = i
        End If
    ElseIf SSTab1.Tab = 8 Then
        Set Cols = TDBGrid_Contract.Columns
        i = TDBGrid_Contract.Col
        TDBGrid_Contract.HoldFields
        
        rsContract.Filter = getFilter()
        TDBGrid_Contract.Col = i
        TDBGrid_Contract.EditActive = True
        
        TDBGrid_Contract.SelStart = Len(TDBGrid_Contract.Columns(i).FilterText)
        If TDBGrid_Contract.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Contract.Col = i
        End If
    End If
    Exit Sub
    
Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub cmdNew_Click(Index As Integer)
    If SSTab1.Tab = 0 Then
        If TDBCombo_company.Text = "" Then
            MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
        
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
    Else
        If Not (TDBGrid_Employee.ApproxCount > 0 And TDBGrid_Employee.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
    End If
    
    Call new_data
End Sub

Private Sub cmdSave_Click(Index As Integer)
    Call simpan_data
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    Call edit_data
End Sub

Private Sub cmdDelete_Click(Index As Integer)
    Call delete_data
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Call cancel_data
End Sub


Private Sub TDBGrid_Employee_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Family_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Education_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Skill_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Job_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Title_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Grade_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Training_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Contract_FilterChange()
    Call grid_filter
End Sub

Private Sub txt_nik_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_employee_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_employee_nick_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_place_of_birth_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_account_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt_family_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_fams_edu_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_fams_employment_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt_edu_majors_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_edu_school_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_edu_city_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_edu_country_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt_skill_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt_job_company_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt_training_subject_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt_contract_no_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtAlamat_LostFocus()
    txt_address.Text = txtAlamat.Text
End Sub


Private Sub txt_job_salary_Validate(Cancel As Boolean)
    If Not Trim(txt_job_salary) = "" Then
        txt_job_salary = FormatNumber(DropAllComma(txt_job_salary))
    End If
End Sub
