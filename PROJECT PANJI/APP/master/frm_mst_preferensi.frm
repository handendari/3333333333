VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_mst_preferensi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER PREFERENSI"
   ClientHeight    =   8490
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   90
      Top             =   7800
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6765
      Left            =   90
      TabIndex        =   32
      Top             =   870
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   11933
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "JHK"
      TabPicture(0)   =   "frm_mst_preferensi.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "fra_entry_jhk"
      Tab(0).Control(2)=   "TDBGrid2"
      Tab(0).Control(3)=   "Label26"
      Tab(0).Control(4)=   "Label4"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "UMK"
      TabPicture(1)   =   "frm_mst_preferensi.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TDBGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra_entry"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frmTombol"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "PERIODE CUTI"
      TabPicture(2)   =   "frm_mst_preferensi.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TDBGrid3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fra_entry_leave"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "GENERAL"
      TabPicture(3)   =   "frm_mst_preferensi.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(1)=   "Frame4"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "TUNJANGAN"
      TabPicture(4)   =   "frm_mst_preferensi.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label28"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "DTPicker_salary"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame12"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.Frame Frame12 
         Caption         =   "Data Control Button"
         Height          =   1275
         Left            =   -74760
         TabIndex        =   125
         Top             =   5340
         Width           =   11085
         Begin prj_panji.vbButton cmdSave_Tunj 
            Height          =   705
            Left            =   9090
            TabIndex        =   22
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Simpan"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":008C
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":00B8
            picn            =   "frm_mst_preferensi.frx":00D6
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -72960
         TabIndex        =   84
         Top             =   3510
         Width           =   7965
         Begin prj_panji.vbButton cmdSave_Gen 
            Height          =   705
            Left            =   6090
            TabIndex        =   85
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Simpan"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":116A
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":1196
            picn            =   "frm_mst_preferensi.frx":11B4
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
      End
      Begin VB.Frame Frame8 
         Height          =   2415
         Left            =   -72960
         TabIndex        =   83
         Top             =   1080
         Width           =   7965
         Begin VB.ComboBox cboSalary 
            Height          =   315
            ItemData        =   "frm_mst_preferensi.frx":2248
            Left            =   4320
            List            =   "frm_mst_preferensi.frx":2252
            TabIndex        =   89
            Text            =   "Combo1"
            Top             =   1290
            Width           =   1575
         End
         Begin VB.TextBox txt_wh_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4320
            TabIndex        =   86
            Top             =   780
            Width           =   1545
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "METODE GAJI"
            Height          =   195
            Left            =   3120
            TabIndex        =   88
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JUMLAH JAM KERJA SEBULAN"
            Height          =   195
            Left            =   1980
            TabIndex        =   87
            Top             =   840
            Width           =   2145
         End
      End
      Begin VB.Frame fra_entry_leave 
         Height          =   2655
         Left            =   -74850
         TabIndex        =   68
         Top             =   2040
         Width           =   11295
         Begin VB.CheckBox chkAllDiv_Leave 
            Caption         =   "SEMUA DIVISI"
            Height          =   225
            Left            =   3360
            TabIndex        =   81
            Top             =   900
            Width           =   1725
         End
         Begin VB.TextBox txt_description_leave 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3360
            MaxLength       =   50
            TabIndex        =   79
            Top             =   1890
            Width           =   5415
         End
         Begin VB.ComboBox cboLeave 
            Height          =   315
            ItemData        =   "frm_mst_preferensi.frx":226E
            Left            =   3360
            List            =   "frm_mst_preferensi.frx":2278
            TabIndex        =   77
            Text            =   "PERIODE KERJA"
            Top             =   1530
            Width           =   1545
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
            TabIndex        =   71
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_company_name_leave 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   4920
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   70
            Top             =   480
            Width           =   3855
         End
         Begin VB.TextBox txt_division_name_leave 
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
            Left            =   4920
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   69
            Top             =   1170
            Width           =   3855
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_company_leave 
            Height          =   375
            Left            =   3360
            OleObjectBlob   =   "frm_mst_preferensi.frx":2292
            TabIndex        =   72
            Top             =   480
            Width           =   1545
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_division_leave 
            Height          =   375
            Left            =   3360
            OleObjectBlob   =   "frm_mst_preferensi.frx":41FE
            TabIndex        =   73
            Top             =   1170
            Width           =   1545
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   1920
            TabIndex        =   80
            Top             =   1890
            Width           =   1230
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TYPE PERIODE CUTI"
            Height          =   195
            Left            =   1695
            TabIndex        =   78
            Top             =   1560
            Width           =   1470
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DIVISI"
            Height          =   195
            Left            =   2685
            TabIndex        =   75
            Top             =   900
            Width           =   465
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PERUSAHAAN"
            Height          =   195
            Left            =   2145
            TabIndex        =   74
            Top             =   510
            Width           =   1005
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74850
         TabIndex        =   62
         Top             =   4800
         Width           =   11295
         Begin prj_panji.vbButton cmdNew_leave 
            Height          =   705
            Left            =   570
            TabIndex        =   63
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Tambah"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":616B
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":6197
            picn            =   "frm_mst_preferensi.frx":61B5
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_leave 
            Height          =   705
            Left            =   1590
            TabIndex        =   64
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Simpan"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":7249
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":7275
            picn            =   "frm_mst_preferensi.frx":7293
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_leave 
            Height          =   705
            Left            =   2610
            TabIndex        =   65
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Ubah"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":8327
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":8353
            picn            =   "frm_mst_preferensi.frx":8371
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_leave 
            Height          =   705
            Left            =   3630
            TabIndex        =   66
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Hapus"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":9405
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":9431
            picn            =   "frm_mst_preferensi.frx":944F
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_leave 
            Height          =   705
            Left            =   4650
            TabIndex        =   67
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Batal"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":A4E3
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":A50F
            picn            =   "frm_mst_preferensi.frx":A52D
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74850
         TabIndex        =   43
         Top             =   4800
         Width           =   11295
         Begin prj_panji.vbButton cmdNew_JHK 
            Height          =   705
            Left            =   570
            TabIndex        =   44
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Tambah"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":B5C1
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":B5ED
            picn            =   "frm_mst_preferensi.frx":B60B
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_JHK 
            Height          =   705
            Left            =   1590
            TabIndex        =   45
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Simpan"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":C69F
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":C6CB
            picn            =   "frm_mst_preferensi.frx":C6E9
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_JHK 
            Height          =   705
            Left            =   2610
            TabIndex        =   46
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Ubah"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":D77D
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":D7A9
            picn            =   "frm_mst_preferensi.frx":D7C7
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_JHK 
            Height          =   705
            Left            =   3630
            TabIndex        =   47
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Hapus"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":E85B
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":E887
            picn            =   "frm_mst_preferensi.frx":E8A5
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_JHK 
            Height          =   705
            Left            =   4650
            TabIndex        =   48
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Batal"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":F939
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":F965
            picn            =   "frm_mst_preferensi.frx":F983
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
      End
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   150
         TabIndex        =   38
         Top             =   4830
         Width           =   11295
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Left            =   570
            TabIndex        =   27
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Tambah"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":10A17
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":10A43
            picn            =   "frm_mst_preferensi.frx":10A61
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave 
            Height          =   705
            Left            =   1590
            TabIndex        =   28
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Simpan"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":11AF5
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":11B21
            picn            =   "frm_mst_preferensi.frx":11B3F
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit 
            Height          =   705
            Left            =   2610
            TabIndex        =   29
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Ubah"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":12BD3
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":12BFF
            picn            =   "frm_mst_preferensi.frx":12C1D
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete 
            Height          =   705
            Left            =   3630
            TabIndex        =   30
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Hapus"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":13CB1
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":13CDD
            picn            =   "frm_mst_preferensi.frx":13CFB
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel 
            Height          =   705
            Left            =   4650
            TabIndex        =   31
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Batal"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":14D8F
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":14DBB
            picn            =   "frm_mst_preferensi.frx":14DD9
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
      End
      Begin VB.Frame fra_entry 
         Height          =   2655
         Left            =   150
         TabIndex        =   33
         Top             =   2070
         Width           =   11295
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
            TabIndex        =   34
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_umk_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   25
            Top             =   1140
            Width           =   3495
         End
         Begin VB.TextBox txt_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   26
            Top             =   1500
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker DTPicker_umk 
            Height          =   315
            Left            =   4800
            TabIndex        =   24
            Top             =   750
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   95223811
            CurrentDate     =   41213
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NILAI UMK*"
            Height          =   195
            Left            =   3360
            TabIndex        =   37
            Top             =   1140
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TANGGAL*"
            Height          =   195
            Left            =   3360
            TabIndex        =   36
            Top             =   780
            Width           =   1230
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   3360
            TabIndex        =   35
            Top             =   1500
            Width           =   1230
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4245
         Left            =   150
         TabIndex        =   39
         Top             =   480
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "TANGGAL"
         Columns(0).DataField=   "umk_date"
         Columns(0).NumberFormat=   "yyyy-MM-dd"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "UMK"
         Columns(1).DataField=   "umk_value"
         Columns(1).NumberFormat=   "Standard"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=5874"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5794"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=514"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=10292"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=10213"
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
         Caption         =   "DAFTAR UMK"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=1"
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
      Begin VB.Frame fra_entry_jhk 
         Height          =   2655
         Left            =   -74850
         TabIndex        =   49
         Top             =   2040
         Width           =   11295
         Begin VB.CheckBox chkAllDiv_JHK 
            Caption         =   "SEMUA DIVISI"
            Height          =   225
            Left            =   3060
            TabIndex        =   82
            Top             =   990
            Width           =   1725
         End
         Begin VB.TextBox txt_description_jhk 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   57
            Top             =   1980
            Width           =   5415
         End
         Begin VB.TextBox txt_value_jhk 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            TabIndex        =   56
            Top             =   1620
            Width           =   1545
         End
         Begin VB.TextBox txt_division_name_jhk 
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
            Left            =   4620
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   53
            Top             =   1260
            Width           =   3855
         End
         Begin VB.TextBox txt_company_name_jhk 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   4620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   52
            Top             =   570
            Width           =   3855
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
            TabIndex        =   50
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_company_jhk 
            Height          =   375
            Left            =   3060
            OleObjectBlob   =   "frm_mst_preferensi.frx":15E6D
            TabIndex        =   54
            Top             =   570
            Width           =   1545
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_division_jhk 
            Height          =   375
            Left            =   3060
            OleObjectBlob   =   "frm_mst_preferensi.frx":17DD7
            TabIndex        =   55
            Top             =   1260
            Width           =   1545
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PERUSAHAAN"
            Height          =   195
            Left            =   1875
            TabIndex        =   61
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DIVISI"
            Height          =   195
            Left            =   2415
            TabIndex        =   60
            Top             =   990
            Width           =   465
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "J H K"
            Height          =   195
            Left            =   2490
            TabIndex        =   59
            Top             =   1650
            Width           =   360
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   1620
            TabIndex        =   58
            Top             =   1980
            Width           =   1230
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   4245
         Left            =   -74850
         TabIndex        =   51
         Top             =   450
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE PERUSAHAAN"
         Columns(0).DataField=   "company_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "PERUSAHAAN"
         Columns(1).DataField=   "company_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "FLAG DIV"
         Columns(2).DataField=   "flag_division"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "KODE DIV."
         Columns(3).DataField=   "division_code"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DIVISI"
         Columns(4).DataField=   "division_name"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "J H K"
         Columns(5).DataField=   "jhk_value"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=4048"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3969"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2963"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2884"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=3493"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=3413"
         Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=2408"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2328"
         Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=8916"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=8837"
         Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
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
         Caption         =   "DAFTAR JHK"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=70,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
         Height          =   4245
         Left            =   -74850
         TabIndex        =   76
         Top             =   450
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE PERUSAHAAN"
         Columns(0).DataField=   "company_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "PERUSAHAAN"
         Columns(1).DataField=   "company_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "FLAG DIV"
         Columns(2).DataField=   "flag_division"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "KODE DIV."
         Columns(3).DataField=   "division_code"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DIVISI"
         Columns(4).DataField=   "division_name"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "LEAVE TYPE"
         Columns(5).DataField=   "leave_type"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "TIPE"
         Columns(6).DataField=   "type"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "KETERANGAN"
         Columns(7).DataField=   "description"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=4048"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3969"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2963"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2884"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=3969"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=3889"
         Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(33)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(35)=   "Column(6).Width=2408"
         Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2328"
         Splits(0)._ColumnProps(38)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(40)=   "Column(7).Width=8440"
         Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=8361"
         Splits(0)._ColumnProps(43)=   "Column(7)._ColStyle=516"
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
         Caption         =   "DAFTAR PERIODE CUTI"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=70,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=74,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
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
      Begin MSComCtl2.DTPicker DTPicker_salary 
         Height          =   315
         Left            =   -73050
         TabIndex        =   135
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   95223811
         CurrentDate     =   40987
      End
      Begin VB.Frame Frame2 
         Height          =   4605
         Left            =   -74760
         TabIndex        =   90
         Top             =   720
         Width           =   11085
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   165
            Left            =   300
            TabIndex        =   148
            Top             =   4320
            Visible         =   0   'False
            Width           =   10425
            _ExtentX        =   18389
            _ExtentY        =   291
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
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
            Left            =   7110
            MaxLength       =   10
            TabIndex        =   20
            Top             =   3450
            Width           =   1035
         End
         Begin VB.TextBox txt_pph21_allowance_to 
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
            Left            =   8370
            MaxLength       =   10
            TabIndex        =   21
            Top             =   3450
            Width           =   1065
         End
         Begin VB.TextBox txt_shift3_allowance_to 
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
            Left            =   8370
            MaxLength       =   10
            TabIndex        =   19
            Top             =   3090
            Width           =   1065
         End
         Begin VB.TextBox txt_shift2_allowance_to 
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
            Left            =   8370
            MaxLength       =   10
            TabIndex        =   17
            Top             =   2730
            Width           =   1065
         End
         Begin VB.TextBox txt_family_allowance_to 
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
            Left            =   8370
            MaxLength       =   10
            TabIndex        =   13
            Top             =   1560
            Width           =   1065
         End
         Begin VB.TextBox txt_meal_allowance_to 
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
            Left            =   8370
            MaxLength       =   10
            TabIndex        =   15
            Top             =   2370
            Width           =   1065
         End
         Begin VB.TextBox txt_performance_allowance_to 
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
            Left            =   8370
            MaxLength       =   10
            TabIndex        =   11
            Top             =   750
            Width           =   1065
         End
         Begin VB.TextBox txt_title_allowance_to 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   9
            Top             =   3570
            Width           =   1065
         End
         Begin VB.TextBox txt_incentive_allowance_to 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   7
            Top             =   2760
            Width           =   1065
         End
         Begin VB.TextBox txt_presence_allowance_to 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   5
            Top             =   1950
            Width           =   1065
         End
         Begin VB.TextBox txt_main_salary_sunday_to 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   3
            Top             =   1140
            Width           =   1065
         End
         Begin VB.TextBox txt_main_salary_to 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   1
            Top             =   780
            Width           =   1065
         End
         Begin VB.CheckBox chk_pph21 
            Caption         =   "Apply to All"
            Height          =   255
            Left            =   9480
            TabIndex        =   134
            Top             =   3480
            Width           =   1545
         End
         Begin VB.CheckBox chk_basic_sunday 
            Caption         =   "Apply to All"
            Height          =   255
            Left            =   4140
            TabIndex        =   132
            Top             =   1170
            Width           =   1755
         End
         Begin VB.CheckBox chk_basic 
            Caption         =   "Apply to All"
            Height          =   255
            Left            =   4140
            TabIndex        =   131
            Top             =   630
            Width           =   1755
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
            Left            =   1725
            MaxLength       =   10
            TabIndex        =   0
            Top             =   780
            Width           =   1035
         End
         Begin VB.Frame Frame13 
            Height          =   525
            Left            =   1725
            TabIndex        =   126
            Top             =   240
            Width           =   2355
            Begin VB.OptionButton opt_monthly_basic 
               Caption         =   "BULANAN"
               Height          =   225
               Left            =   1110
               TabIndex        =   128
               Top             =   210
               Width           =   1125
            End
            Begin VB.OptionButton opt_daily_basic 
               Caption         =   "HARIAN"
               Height          =   225
               Left            =   90
               TabIndex        =   127
               Top             =   210
               Value           =   -1  'True
               Width           =   945
            End
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
            Left            =   1725
            MaxLength       =   10
            TabIndex        =   2
            Top             =   1140
            Width           =   1035
         End
         Begin VB.CheckBox chk_shift3 
            Caption         =   "Apply to All"
            Height          =   255
            Left            =   9480
            TabIndex        =   124
            Top             =   3120
            Width           =   1575
         End
         Begin VB.CheckBox chk_shift2 
            Caption         =   "Apply to All"
            Height          =   255
            Left            =   9480
            TabIndex        =   123
            Top             =   2760
            Width           =   1575
         End
         Begin VB.CheckBox chk_trans_makan 
            Caption         =   "Apply to All"
            Height          =   255
            Left            =   9480
            TabIndex        =   122
            Top             =   2250
            Width           =   1575
         End
         Begin VB.CheckBox chk_prestasi 
            Caption         =   "Apply to All"
            Height          =   255
            Left            =   9510
            TabIndex        =   120
            Top             =   630
            Width           =   1365
         End
         Begin VB.CheckBox chk_jabatan 
            Caption         =   "Apply to All"
            Height          =   255
            Left            =   4140
            TabIndex        =   119
            Top             =   3450
            Width           =   1755
         End
         Begin VB.CheckBox chk_hadir 
            Caption         =   "Apply to All"
            Height          =   255
            Left            =   4140
            TabIndex        =   117
            Top             =   1830
            Width           =   1755
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
            Left            =   7110
            MaxLength       =   10
            TabIndex        =   12
            Top             =   1560
            Width           =   1035
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
            Left            =   7110
            MaxLength       =   10
            TabIndex        =   10
            Top             =   750
            Width           =   1035
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
            Left            =   1740
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1950
            Width           =   1035
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
            Left            =   1740
            MaxLength       =   10
            TabIndex        =   8
            Top             =   3570
            Width           =   1035
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
            Left            =   7110
            MaxLength       =   10
            TabIndex        =   18
            Top             =   3090
            Width           =   1035
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
            Left            =   7110
            MaxLength       =   10
            TabIndex        =   14
            Top             =   2370
            Width           =   1035
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
            Left            =   7110
            MaxLength       =   10
            TabIndex        =   16
            Top             =   2730
            Width           =   1035
         End
         Begin VB.Frame Frame5 
            Height          =   525
            Left            =   1740
            TabIndex        =   94
            Top             =   1410
            Width           =   2355
            Begin VB.OptionButton opt_monthly_presence 
               Caption         =   "BULANAN"
               Height          =   225
               Left            =   1110
               TabIndex        =   95
               Top             =   210
               Width           =   1125
            End
            Begin VB.OptionButton opt_daily_presence 
               Caption         =   "HARIAN"
               Height          =   225
               Left            =   90
               TabIndex        =   96
               Top             =   210
               Value           =   -1  'True
               Width           =   1065
            End
         End
         Begin VB.Frame Frame6 
            Height          =   525
            Left            =   7110
            TabIndex        =   91
            Top             =   1020
            Width           =   2355
            Begin VB.OptionButton opt_monthly_family 
               Caption         =   "BULANAN"
               Height          =   225
               Left            =   1110
               TabIndex        =   92
               Top             =   210
               Value           =   -1  'True
               Width           =   1125
            End
            Begin VB.OptionButton opt_daily_family 
               Caption         =   "HARIAN"
               Height          =   225
               Left            =   90
               TabIndex        =   93
               Top             =   210
               Width           =   1065
            End
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
            Left            =   1740
            MaxLength       =   10
            TabIndex        =   6
            Top             =   2760
            Width           =   1035
         End
         Begin VB.Frame Frame11 
            Height          =   525
            Left            =   1740
            TabIndex        =   106
            Top             =   2220
            Width           =   2355
            Begin VB.OptionButton opt_monthly_insentif 
               Caption         =   "BULANAN"
               Height          =   225
               Left            =   1110
               TabIndex        =   107
               Top             =   210
               Value           =   -1  'True
               Width           =   1125
            End
            Begin VB.OptionButton opt_daily_insentif 
               Caption         =   "HARIAN"
               Height          =   225
               Left            =   90
               TabIndex        =   108
               Top             =   210
               Width           =   1065
            End
         End
         Begin VB.Frame Frame7 
            Height          =   525
            Left            =   1740
            TabIndex        =   97
            Top             =   3030
            Width           =   2355
            Begin VB.OptionButton opt_monthly_title 
               Caption         =   "BULANAN"
               Height          =   225
               Left            =   1110
               TabIndex        =   98
               Top             =   210
               Value           =   -1  'True
               Width           =   1125
            End
            Begin VB.OptionButton opt_daily_title 
               Caption         =   "HARIAN"
               Height          =   225
               Left            =   90
               TabIndex        =   99
               Top             =   210
               Width           =   1065
            End
         End
         Begin VB.Frame Frame9 
            Height          =   525
            Left            =   7110
            TabIndex        =   100
            Top             =   210
            Width           =   2355
            Begin VB.OptionButton opt_monthly_performance 
               Caption         =   "BULANAN"
               Height          =   225
               Left            =   1110
               TabIndex        =   101
               Top             =   210
               Value           =   -1  'True
               Width           =   1125
            End
            Begin VB.OptionButton opt_daily_performance 
               Caption         =   "HARIAN"
               Height          =   225
               Left            =   90
               TabIndex        =   102
               Top             =   210
               Width           =   1065
            End
         End
         Begin VB.Frame Frame10 
            Height          =   525
            Left            =   7110
            TabIndex        =   103
            Top             =   1830
            Width           =   2355
            Begin VB.OptionButton opt_monthly_meal 
               Caption         =   "BULANAN"
               Height          =   225
               Left            =   1110
               TabIndex        =   105
               Top             =   210
               Width           =   1125
            End
            Begin VB.OptionButton opt_daily_meal 
               Caption         =   "HARIAN"
               Height          =   225
               Left            =   90
               TabIndex        =   104
               Top             =   210
               Value           =   -1  'True
               Width           =   1005
            End
         End
         Begin VB.CheckBox chk_insentif 
            Caption         =   "Apply to All"
            Height          =   255
            Left            =   4140
            TabIndex        =   118
            Top             =   2640
            Width           =   1755
         End
         Begin VB.CheckBox chk_keluarga 
            Caption         =   "Apply to All"
            Height          =   255
            Left            =   9480
            TabIndex        =   121
            Top             =   1410
            Width           =   1575
         End
         Begin VB.Label lbl_ket 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   420
            TabIndex        =   149
            Top             =   4080
            Width           =   10245
         End
         Begin VB.Label Label39 
            Caption         =   "Ke"
            Height          =   195
            Left            =   8160
            TabIndex        =   147
            Top             =   3510
            Width           =   165
         End
         Begin VB.Label Label38 
            Caption         =   "Ke"
            Height          =   195
            Left            =   8160
            TabIndex        =   146
            Top             =   3150
            Width           =   165
         End
         Begin VB.Label Label37 
            Caption         =   "Ke"
            Height          =   195
            Left            =   8160
            TabIndex        =   145
            Top             =   2760
            Width           =   165
         End
         Begin VB.Label Label36 
            Caption         =   "Ke"
            Height          =   195
            Left            =   8160
            TabIndex        =   144
            Top             =   2430
            Width           =   165
         End
         Begin VB.Label Label35 
            Caption         =   "Ke"
            Height          =   195
            Left            =   8160
            TabIndex        =   143
            Top             =   1620
            Width           =   165
         End
         Begin VB.Label Label34 
            Caption         =   "Ke"
            Height          =   195
            Left            =   8160
            TabIndex        =   142
            Top             =   810
            Width           =   165
         End
         Begin VB.Label Label33 
            Caption         =   "Ke"
            Height          =   195
            Left            =   2790
            TabIndex        =   141
            Top             =   3630
            Width           =   165
         End
         Begin VB.Label Label32 
            Caption         =   "Ke"
            Height          =   195
            Left            =   2790
            TabIndex        =   140
            Top             =   2820
            Width           =   165
         End
         Begin VB.Label Label31 
            Caption         =   "Ke"
            Height          =   195
            Left            =   2790
            TabIndex        =   139
            Top             =   2010
            Width           =   165
         End
         Begin VB.Label Label30 
            Caption         =   "Ke"
            Height          =   195
            Left            =   2790
            TabIndex        =   138
            Top             =   1200
            Width           =   165
         End
         Begin VB.Label Label29 
            Caption         =   "Ke"
            Height          =   195
            Left            =   2790
            TabIndex        =   137
            Top             =   840
            Width           =   165
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PPh 21"
            Height          =   195
            Left            =   6330
            TabIndex        =   133
            Top             =   3510
            Width           =   555
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "GAJI POKOK"
            Height          =   195
            Left            =   615
            TabIndex        =   130
            Top             =   660
            Width           =   900
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "GAPOK (MINGGU)"
            Height          =   195
            Left            =   255
            TabIndex        =   129
            Top             =   1170
            Width           =   1275
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KELUARGA"
            Height          =   195
            Left            =   5970
            TabIndex        =   116
            Top             =   1470
            Width           =   900
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PRESTASI"
            Height          =   195
            Left            =   6075
            TabIndex        =   115
            Top             =   630
            Width           =   795
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JABATAN"
            Height          =   195
            Left            =   450
            TabIndex        =   113
            Top             =   3450
            Width           =   1050
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SHIFT 3"
            Height          =   195
            Left            =   5820
            TabIndex        =   112
            Top             =   3150
            Width           =   1050
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TRNS/MAKAN"
            Height          =   195
            Left            =   5835
            TabIndex        =   111
            Top             =   2250
            Width           =   1035
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SHIFT 2"
            Height          =   195
            Left            =   5970
            TabIndex        =   110
            Top             =   2760
            Width           =   900
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KEHADIRAN"
            Height          =   195
            Left            =   405
            TabIndex        =   114
            Top             =   1830
            Width           =   1095
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "INSENTIF KERJA"
            Height          =   195
            Left            =   225
            TabIndex        =   109
            Top             =   2670
            Width           =   1290
         End
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TANGGAL EFEKTIF"
         Height          =   195
         Left            =   -74520
         TabIndex        =   136
         Top             =   450
         Width           =   1335
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PERUSAHAAN*"
         Height          =   195
         Left            =   -72420
         TabIndex        =   42
         Top             =   2100
         Width           =   1170
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DIVISI"
         Height          =   195
         Left            =   -71715
         TabIndex        =   41
         Top             =   2430
         Width           =   465
      End
   End
   Begin prj_panji.vbButton cmdExit 
      Height          =   705
      Left            =   10440
      TabIndex        =   40
      Top             =   7710
      Width           =   945
      _extentx        =   1667
      _extenty        =   1244
      btype           =   14
      tx              =   "&Keluar"
      enab            =   -1
      font            =   "frm_mst_preferensi.frx":19D42
      coltype         =   1
      focusr          =   -1
      bcol            =   15790320
      bcolo           =   15790320
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frm_mst_preferensi.frx":19D6E
      picn            =   "frm_mst_preferensi.frx":19D8C
      umcol           =   -1
      soft            =   0
      picpos          =   2
      ngrey           =   0
      fx              =   0
      hand            =   0
      check           =   0
      value           =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER PREFERENSI"
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
      TabIndex        =   23
      Top             =   150
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_preferensi.frx":1AE20
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11790
   End
End
Attribute VB_Name = "frm_mst_preferensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUMK As New ADODB.Recordset
Dim rsJHK As New ADODB.Recordset
Dim rsLeave As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsDivision As New ADODB.Recordset
Dim rsGeneral As New ADODB.Recordset
Dim rsPrefTunj As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Public public_int_mode As Integer

Dim vFlagBasic As Integer, vBasic As String
Dim vBasicSunday As String
Dim vFlagHadir As Integer, vHadir As String
Dim vFlagInsentif As Integer, vInsentif As String
Dim vFlagJabatan As Integer, vJabatan As String
Dim vFlagPrestasi As Integer, vPrestasi As String
Dim vFlagKeluarga As Integer, vKeluarga As String
Dim vFlagTrans As Integer, vTrans As String
Dim vShift2 As String, vShift3 As String
Dim vPPh21 As String

Private Function check_validate_exist_new() As Boolean
Dim str_sql As String
    check_validate_exist_new = False
    
    If SSTab1.Tab = 0 Then
        If rs.State Then rs.Close
        str_sql = "select count(company_code) as rec_count from m_pref_jhk " & _
                    "where company_code = '" & TDBCombo_company_jhk.Text & "' and division_code = '" & TDBCombo_division_jhk.Text & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
        
        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            rs.Close
            Exit Function
        End If
        
        rs.Close
    ElseIf SSTab1.Tab = 1 Then
        If rs.State Then rs.Close
        str_sql = "select count(umk_date) as rec_count from m_pref_umk " & _
                    "where date(umk_date) = '" & Format(DTPicker_umk.Value, "yyyy-MM-dd") & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
        
        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            rs.Close
            Exit Function
        End If
        
        rs.Close
    ElseIf SSTab1.Tab = 2 Then
        If rs.State Then rs.Close
        str_sql = "select count(company_code) as rec_count from m_pref_leave " & _
                    "where company_code = '" & TDBCombo_company_leave.Text & "' and division_code = '" & TDBCombo_division_leave.Text & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
        
        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            rs.Close
            Exit Function
        End If
        
        rs.Close
    End If
End Function

Private Sub check_invalid()
    If SSTab1.Tab = 0 Then
        MsgBox "Data Sudah Ada...", vbCritical, headerMSG
        TDBCombo_company_jhk.Text = ""
        If TDBCombo_company_jhk.Enabled = True Then TDBCombo_company_jhk.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        MsgBox "Data Sudah Ada...", vbCritical, headerMSG
        DTPicker_umk.Value = Now
        If DTPicker_umk.Enabled = True Then txt_umk_value.SetFocus
    ElseIf SSTab1.Tab = 2 Then
        MsgBox "Data Sudah Ada...", vbCritical, headerMSG
        TDBCombo_company_leave.Text = ""
        If TDBCombo_company_leave.Enabled = True Then TDBCombo_company_leave.SetFocus
    End If
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If SSTab1.Tab = 0 Then
        If Not TDBCombo_company_jhk.Text = rsJHK.Fields("company_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        If Not DTPicker_umk.Value = rsUMK.Fields("umk_date").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        If Not TDBCombo_company_leave.Text = rsLeave.Fields("company_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True
    
    If SSTab1.Tab = 0 Then
        'validasi tdb_company
        If Trim(TDBCombo_company_jhk.Text) = "" Then
            MsgBox "Perusahaan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_company_jhk.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If chkAllDiv_JHK.Value = 0 Then
            'validasi tdb_division
            If Trim(TDBCombo_division_jhk.Text) = "" Then
                MsgBox "Divisi Masih Kosong...", vbOKOnly + vbInformation, headerMSG
                TDBCombo_division_jhk.SetFocus
                check_validate_new = False
                Exit Function
            End If
        End If
        
        'validasi jhk_value
        If Trim(txt_value_jhk.Text) = "" Then
            MsgBox "JHK Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_value_jhk.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        'validasi umk
        If Trim(txt_umk_value) = "" Then
            MsgBox "Nilai UMK Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_umk_value.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        'validasi tdb_company
        If Trim(TDBCombo_company_leave.Text) = "" Then
            MsgBox "Perusahaan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_company_leave.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If chkAllDiv_Leave.Value = 0 Then
            'validasi tdb_division
            If Trim(TDBCombo_division_leave.Text) = "" Then
                MsgBox "Divisi Masih Kosong...", vbOKOnly + vbInformation, headerMSG
                TDBCombo_division_leave.SetFocus
                check_validate_new = False
                Exit Function
            End If
        End If
    End If
End Function

Private Sub cancel_data()
    int_mode = 0
    Call load_mode
End Sub

Private Sub delete_data()
    Dim i As Integer
    Dim vWhere As String
    Dim vCompany As String
    Dim vDivision As String
    Dim vFlagDivision As Integer
    
    If SSTab1.Tab = 0 Then
        If Not (TDBGrid2.ApproxCount > 0 And TDBGrid2.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid2.Columns("company_name").Value & "' - '" & TDBGrid2.Columns("division_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        vCompany = TDBGrid2.Columns("company_code").Value
        vDivision = TDBGrid2.Columns("division_code").Value
        
        CnG.Execute "delete from m_pref_jhk " & _
            "where company_code = '" & vCompany & "' and division_code = '" & vDivision & "'"
        CnG.CommitTrans
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid1.Columns("umk_date").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_pref_umk where date(umk_date) = '" _
            & Format(TDBGrid1.Columns("umk_date").Value, "yyyy-MM-dd") & "'"
        CnG.CommitTrans
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid3.ApproxCount > 0 And TDBGrid3.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid3.Columns("company_name").Value & "' - '" & TDBGrid3.Columns("division_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        vCompany = TDBGrid3.Columns("company_code").Value
        vDivision = TDBGrid3.Columns("division_code").Value
        
        CnG.Execute "delete from m_pref_leave " & _
            "where company_code = '" & vCompany & "' and division_code = '" & vDivision & "'"
        CnG.CommitTrans
    End If
    
    Call load_data
    int_mode = 0
    Call load_mode
End Sub

Public Sub set_edit_data()
On Error GoTo Err
    vSetData = 1
    
    If SSTab1.Tab = 0 Then
        If Not (TDBGrid2.ApproxCount > 0 And TDBGrid2.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsJHK
            TDBCombo_company_jhk.Text = .Fields("company_code").Value
            txt_company_name_jhk.Text = .Fields("company_name").Value
            chkAllDiv_JHK.Value = IIf(IsNull(.Fields("flag_division").Value), 0, .Fields("flag_division").Value)
            TDBCombo_division_jhk.Text = IIf(IsNull(.Fields("division_code").Value), "", .Fields("division_code").Value)
            txt_division_name_jhk.Text = IIf(IsNull(.Fields("division_name").Value), "", .Fields("division_name").Value)
            txt_value_jhk.Text = .Fields("jhk_value").Value
            txt_description_jhk.Text = .Fields("description").Value
        End With
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsUMK
            DTPicker_umk.Value = .Fields("umk_date").Value
            txt_umk_value = FormatNumber(.Fields("umk_value").Value)
            txt_description = .Fields("description").Value
        End With
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid3.ApproxCount > 0 And TDBGrid3.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsLeave
            TDBCombo_company_leave.Text = .Fields("company_code").Value
            txt_company_name_leave.Text = .Fields("company_name").Value
            chkAllDiv_Leave.Value = IIf(IsNull(.Fields("flag_division").Value), 0, .Fields("flag_division").Value)
            TDBCombo_division_leave.Text = IIf(IsNull(.Fields("division_code").Value), "", .Fields("division_code").Value)
            txt_division_name_leave.Text = IIf(IsNull(.Fields("division_name").Value), "", .Fields("division_name").Value)
            cboLeave.ListIndex = .Fields("leave_type").Value
            txt_description_leave.Text = .Fields("description").Value
        End With
    ElseIf SSTab1.Tab = 3 Then
        If rsGeneral.RecordCount > 0 Then
            With rsGeneral
                txt_wh_value.Text = IIf(IsNull(.Fields("wh_value").Value), 0, .Fields("wh_value").Value)
                cboSalary.ListIndex = IIf(IsNull(.Fields("salary_type").Value), 0, .Fields("salary_type").Value)
            End With
        End If
    ElseIf SSTab1.Tab = 4 Then
    Dim v_flag_basic As Integer
    Dim v_flag_presence As Integer
    Dim v_flag_incentive As Integer
    Dim v_flag_title As Integer
    Dim v_flag_performance As Integer
    Dim v_flag_family As Integer
    Dim v_flag_meal As Integer

        If rsPrefTunj.RecordCount > 0 Then
            With rsPrefTunj
                DTPicker_salary.Value = Now
                
                v_flag_basic = IIf(IsNull(.Fields("flag_basic").Value), 0, .Fields("flag_basic").Value)
                opt_daily_basic.Value = IIf(v_flag_basic = 0, True, False)
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
                
                v_flag_family = IIf(IsNull(.Fields("flag_performance").Value), 0, .Fields("flag_performance").Value)
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
                                
                                
                vFlagBasic = IIf(v_flag_basic = 0, True, False)
                txt_main_salary_to.Text = FormatNumber(IIf(IsNull(.Fields("basic_salary").Value), 0, .Fields("basic_salary").Value))
                txt_main_salary_sunday_to.Text = FormatNumber(IIf(IsNull(.Fields("basic_salary_sunday").Value), 0, .Fields("basic_salary_sunday").Value))
                                
                vFlagHadir = v_flag_presence
                txt_presence_allowance_to.Text = FormatNumber(IIf(IsNull(.Fields("presence_allowance").Value), 0, .Fields("presence_allowance").Value))
                                
                vFlagInsentif = v_flag_incentive
                txt_incentive_allowance_to.Text = FormatNumber(IIf(IsNull(.Fields("incentive_allowance").Value), 0, .Fields("incentive_allowance").Value))
                                
                vFlagJabatan = v_flag_title
                txt_title_allowance_to.Text = FormatNumber(IIf(IsNull(.Fields("title_allowance").Value), 0, .Fields("title_allowance").Value))
                                
                vFlagPrestasi = v_flag_performance
                txt_performance_allowance_to.Text = FormatNumber(IIf(IsNull(.Fields("performance_allowance").Value), 0, .Fields("performance_allowance").Value))
                                
                vFlagKeluarga = v_flag_family
                txt_family_allowance_to.Text = FormatNumber(IIf(IsNull(.Fields("family_allowance").Value), 0, .Fields("family_allowance").Value))
                                
                vFlagTrans = v_flag_meal
                txt_meal_allowance_to.Text = FormatNumber(IIf(IsNull(.Fields("meal_allowance").Value), 0, .Fields("meal_allowance").Value))
                
                txt_shift2_allowance_to.Text = FormatNumber(IIf(IsNull(.Fields("shift2_allowance").Value), 0, .Fields("shift2_allowance").Value))
                txt_shift3_allowance_to.Text = FormatNumber(IIf(IsNull(.Fields("shift3_allowance").Value), 0, .Fields("shift3_allowance").Value))
                txt_pph21_allowance_to.Text = FormatNumber(IIf(IsNull(.Fields("pph21_allowance").Value), 0, .Fields("pph21_allowance").Value))
            End With
        End If
    End If
    
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_data()
    int_mode = 2
    Call load_mode
End Sub

Private Sub chkAllDiv_JHK_Click()
    If chkAllDiv_JHK.Value Then
        TDBCombo_division_jhk.Enabled = False
    Else
        TDBCombo_division_jhk.Enabled = True
    End If
End Sub

Private Sub chkAllDiv_Leave_Click()
    If chkAllDiv_Leave.Value Then
        TDBCombo_division_leave.Enabled = False
    Else
        TDBCombo_division_leave.Enabled = True
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
        If chkAllDiv_JHK.Value Then
            SQL = "SELECT division_code FROM m_division WHERE company_code = '" & TDBCombo_company_jhk.Text & "'"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                While Not rs.EOF
                    SQL = "DELETE FROM m_pref_jhk WHERE company_code = '" & TDBCombo_company_jhk.Text & "' and  division_code = '" & rs!division_code & "'"
                    CnG.Execute SQL
                    
                    SQL = "INSERT INTO m_pref_jhk(company_code,flag_division,division_code,jhk_value,description,entry_date,entry_user) " _
                            & "VALUES " _
                            & "('" & TDBCombo_company_jhk.Text & "',0," _
                            & "'" & rs!division_code & "'," _
                            & "'" & txt_value_jhk.Text & "','" & Trim(txt_description_jhk.Text) & "',now(),'" & LOGIN_NAME & "')"
                    CnG.Execute SQL
                    rs.MoveNext
                Wend
            End If
            rs.Close
        Else
            SQL = "INSERT INTO m_pref_jhk(company_code,flag_division,division_code,jhk_value,description,entry_date,entry_user) " _
                    & "VALUES " _
                    & "('" & TDBCombo_company_jhk.Text & "',0," _
                    & "'" & TDBCombo_division_jhk.Text & "'," _
                    & "'" & txt_value_jhk.Text & "','" & Trim(txt_description_jhk.Text) & "',now(),'" & LOGIN_NAME & "')"
            CnG.Execute SQL
        End If
    ElseIf SSTab1.Tab = 1 Then
        SQL = "INSERT INTO m_pref_umk(umk_date,umk_value,description) " _
                & "VALUES " _
                & "('" & Format(DTPicker_umk, "yyyy-MM-dd") & "'," _
                & "'" & Val(DropAllComma(txt_umk_value.Text)) & "','" & Trim(txt_description.Text) & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        If chkAllDiv_Leave.Value Then
            SQL = "SELECT division_code FROM m_division WHERE company_code = '" & TDBCombo_company_leave.Text & "'"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                While Not rs.EOF
                    SQL = "DELETE FROM m_pref_leave WHERE company_code = '" & TDBCombo_company_leave.Text & "' and division_code = '" & rs!division_code & "'"
                    CnG.Execute SQL
                    
                    SQL = "INSERT INTO m_pref_leave(company_code,flag_division,division_code,leave_type,description,entry_date,entry_user) " _
                            & "VALUES " _
                            & "('" & TDBCombo_company_leave.Text & "',0," _
                            & "'" & rs!division_code & "'," _
                            & "'" & cboLeave.ListIndex & "','" & Trim(txt_description_leave.Text) & "',now(),'" & LOGIN_NAME & "')"
                    CnG.Execute SQL
                    rs.MoveNext
                Wend
            End If
            rs.Close
        Else
            SQL = "INSERT INTO m_pref_leave(company_code,flag_division,division_code,leave_type,description,entry_date,entry_user) " _
                    & "VALUES " _
                    & "('" & TDBCombo_company_leave.Text & "',0," _
                    & "'" & TDBCombo_division_leave.Text & "'," _
                    & "'" & cboLeave.ListIndex & "','" & Trim(txt_description_leave.Text) & "',now(),'" & LOGIN_NAME & "')"
            CnG.Execute SQL
        End If
    ElseIf SSTab1.Tab = 3 Then
        SQL = "DELETE FROM m_pref_gen"
        CnG.Execute SQL
            
        SQL = "INSERT INTO m_pref_gen(wh_value,salary_type) " _
                & "VALUES " _
                & "('" & txt_wh_value.Text & "'," _
                & "'" & cboSalary.ListIndex & "')"
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
        SQL = "UPDATE m_pref_jhk SET jhk_value = '" & txt_value_jhk.Text & "'," _
                & "description = '" & Trim(txt_description_jhk.Text) & "' " _
                & "WHERE company_code = '" & TDBCombo_company_jhk.Text & "' and division_code = '" & TDBCombo_division_jhk.Text & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "UPDATE m_pref_umk SET umk_value = '" & Val(DropAllComma(txt_umk_value)) & "'," _
                & "description = '" & Trim(txt_description.Text) & "' " _
                & "WHERE date(umk_date) = '" & Format(DTPicker_umk.Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "UPDATE m_pref_leave SET leave_type = '" & cboLeave.ListIndex & "'," _
                & "description = '" & Trim(txt_description_leave.Text) & "' " _
                & "WHERE company_code = '" & TDBCombo_company_leave.Text & "' and division_code = '" & TDBCombo_division_leave.Text & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 4 Then
        SQL = "UPDATE m_pref_tunj SET " & _
            "flag_basic = '" & IIf(opt_daily_basic.Value, 0, 1) & "'," & _
            "basic_salary = '" & Val(DropAllComma(txt_main_salary_to.Text)) & "'," & _
            "basic_salary_sunday = '" & Val(DropAllComma(txt_main_salary_sunday_to.Text)) & "'," & _
            "flag_presence = '" & IIf(opt_daily_presence.Value, 0, 1) & "'," & _
            "presence_allowance = '" & Val(DropAllComma(txt_presence_allowance_to.Text)) & "'," & _
            "flag_incentive = '" & IIf(opt_daily_insentif.Value, 0, 1) & "'," & _
            "incentive_allowance = '" & Val(DropAllComma(txt_incentive_allowance_to.Text)) & "'," & _
            "flag_title = '" & IIf(opt_daily_title.Value, 0, 1) & "'," & _
            "title_allowance = '" & Val(DropAllComma(txt_title_allowance_to.Text)) & "'," & _
            "flag_performance = '" & IIf(opt_daily_performance.Value, 0, 1) & "'," & _
            "performance_allowance = '" & Val(DropAllComma(txt_performance_allowance_to.Text)) & "'," & _
            "flag_family = '" & IIf(opt_daily_family.Value, 0, 1) & "'," & _
            "family_allowance = '" & Val(DropAllComma(txt_family_allowance_to.Text)) & "'," & _
            "flag_meal = '" & IIf(opt_daily_meal.Value, 0, 1) & "'," & _
            "meal_allowance = '" & Val(DropAllComma(txt_meal_allowance_to.Text)) & "'," & _
            "shift2_allowance = '" & Val(DropAllComma(txt_shift2_allowance_to.Text)) & "'," & _
            "shift3_allowance = '" & Val(DropAllComma(txt_shift3_allowance_to.Text)) & "'," & _
            "pph21_allowance = '" & Val(DropAllComma(txt_pph21_allowance_to.Text)) & "'"
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
        cmdNew_JHK.Enabled = a And blnUser_Add
        cmdSave_JHK.Enabled = b
        cmdEdit_JHK.Enabled = c And blnUser_Edit
        cmdDelete_JHK.Enabled = d And blnUser_Delete
        cmdCancel_JHK.Enabled = e
    ElseIf SSTab1.Tab = 1 Then
        cmdNew.Enabled = a And blnUser_Add
        cmdSave.Enabled = b
        cmdEdit.Enabled = c And blnUser_Edit
        cmdDelete.Enabled = d And blnUser_Delete
        cmdCancel.Enabled = e
    ElseIf SSTab1.Tab = 2 Then
        cmdNew_leave.Enabled = a And blnUser_Add
        cmdSave_leave.Enabled = b
        cmdEdit_leave.Enabled = c And blnUser_Edit
        cmdDelete_leave.Enabled = d And blnUser_Delete
        cmdCancel_leave.Enabled = e
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
        ElseIf TypeOf Ctr Is CheckBox Then
            Ctr.Value = 0
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
    
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        
        If SSTab1.Tab = 0 Then
            fra_entry_jhk.Visible = True
            TDBCombo_company_jhk.Enabled = True
            chkAllDiv_JHK.Enabled = True
            TDBCombo_division_jhk.Enabled = True
            TDBGrid2.Enabled = False
            Call set_new_data
            
            If TDBCombo_company_jhk.Enabled = True Then
                TDBCombo_company_jhk.SetFocus
            End If
        ElseIf SSTab1.Tab = 1 Then
            fra_entry.Visible = True
            DTPicker_umk.Enabled = True
            TDBGrid1.Enabled = False
            Call set_new_data
            
            If DTPicker_umk.Enabled = True Then
                DTPicker_umk.SetFocus
            End If
        ElseIf SSTab1.Tab = 2 Then
            fra_entry_leave.Visible = True
            TDBCombo_company_leave.Enabled = True
            chkAllDiv_Leave.Enabled = True
            TDBCombo_division_leave.Enabled = True
            TDBGrid3.Enabled = False
            Call set_new_data
            
            If TDBCombo_company_leave.Enabled = True Then
                TDBCombo_company_leave.SetFocus
            End If
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        
        If SSTab1.Tab = 0 Then
            fra_entry_jhk.Visible = False
            TDBGrid2.Enabled = True
        ElseIf SSTab1.Tab = 1 Then
            fra_entry.Visible = False
            TDBGrid1.Enabled = True
        ElseIf SSTab1.Tab = 2 Then
            fra_entry_leave.Visible = False
            TDBGrid3.Enabled = True
        End If
        
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        If SSTab1.Tab = 0 Then
            TDBCombo_company_jhk.Enabled = False
            chkAllDiv_JHK.Enabled = False
            TDBCombo_division_jhk.Enabled = False
            fra_entry_jhk.Visible = True
            TDBGrid2.Enabled = False
        ElseIf SSTab1.Tab = 1 Then
            DTPicker_umk.Enabled = False
            fra_entry.Visible = True
            TDBGrid1.Enabled = False
        ElseIf SSTab1.Tab = 2 Then
            TDBCombo_company_leave.Enabled = False
            chkAllDiv_Leave.Enabled = False
            TDBCombo_division_leave.Enabled = False
            fra_entry_leave.Visible = True
            TDBGrid3.Enabled = False
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

Private Sub cmdSave_Gen_Click()
    Call insert_new_data
    
    MsgBox "Save Succesfully...", vbInformation, headerMSG
End Sub

Private Sub cmdSave_Tunj_Click()
Dim i As Integer
On Error GoTo Err
    
    i = MsgBox("Data ini akan disimpan per tanggal efektif '" & _
                    Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'..." & Chr(13) & _
               "Apakah Anda yakin akan melakukan Update?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    Call edit_old_data
    
    CnG.BeginTrans
    DoEvents
    Me.MousePointer = vbHourglass
    
    vBasic = "": vBasicSunday = "": vHadir = "": vInsentif = "": vJabatan = ""
    vPrestasi = "": vKeluarga = "": vTrans = "": vShift2 = "": vShift3 = "": vPPh21 = ""
    '++++++++++++++++++++ Basic Salary +++++++++++++++++++++++
    If chk_basic.Value = 0 Then
        If DropAllComma(txt_main_salary_to.Text) <> DropAllComma(txt_main_salary.Text) Then
            SQL = "SELECT (SELECT employee_code FROM m_salary_standard " & _
                    "WHERE employee_code = a.employee_code " & _
                        "AND date(salary_date) <= '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date desc limit 1) " & _
                  "FROM m_salary_standard a " & _
                  "WHERE basic_salary = " & DropAllComma(txt_main_salary.Text) & " " & _
                  "GROUP BY employee_code"
'            SQL = "SELECT  DISTINCT employee_code FROM m_salary_standard " & _
'                  "WHERE basic_salary = " & DropAllComma(txt_main_salary.Text) & ""
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = 0
                ProgressBar1.Max = rscari.RecordCount
                
                lbl_ket.Visible = True
                lbl_ket.Caption = "Mengupdate Gaji Pokok..."
                
                vBasic = "- Gaji Pokok = " & rscari.RecordCount & " Karyawan"
            
                rscari.MoveFirst
                While Not rscari.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
                    SQL = "SELECT * FROM m_salary_standard " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        SQL = "UPDATE m_salary_standard " & _
                              "SET basic_salary = '" & DropAllComma(txt_main_salary_to.Text) & "' " & _
                              "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & rscari.Fields(0).Value & "'"
                        CnG.Execute SQL
                    Else
                        SQL = "INSERT INTO m_salary_standard " & _
                              "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                                "flag_basic, '" & DropAllComma(txt_main_salary_to.Text) & "'," & _
                                "basic_salary_sunday, flag_presence, presence_allowance," & _
                                "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                                "flag_performance, performance_allowance, flag_family, family_allowance," & _
                                "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                                "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                                "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                              "FROM m_salary_standard " & _
                              "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                                "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                        CnG.Execute SQL
                    End If
                    rs.Close
                    
                rscari.MoveNext
                Wend
            End If
            rscari.Close
            
            ProgressBar1.Visible = False
            lbl_ket.Visible = False
        End If
    Else
        SQL = "SELECT DISTINCT employee_code FROM m_salary_standard"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Max = rscari.RecordCount
            
            lbl_ket.Visible = True
            lbl_ket.Caption = "Mengupdate Gaji Pokok..."
            
            vBasic = "- Gaji Pokok = " & rscari.RecordCount & " Karyawan"
                
            rscari.MoveFirst
            While Not rscari.EOF
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                SQL = "SELECT * FROM m_salary_standard " & _
                      "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                            "AND employee_code = '" & rscari.Fields(0).Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    SQL = "UPDATE m_salary_standard " & _
                          "SET basic_salary = '" & DropAllComma(txt_main_salary_to.Text) & "' " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_salary_standard " & _
                          "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                            "flag_basic, '" & DropAllComma(txt_main_salary_to.Text) & "'," & _
                            "basic_salary_sunday, flag_presence, presence_allowance," & _
                            "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                            "flag_performance, performance_allowance, flag_family, family_allowance," & _
                            "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                            "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                            "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                          "FROM m_salary_standard " & _
                          "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    CnG.Execute SQL
                End If
                rs.Close
                
            rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        ProgressBar1.Visible = False
        lbl_ket.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++ Basic Salary Sunday +++++++++++++++++++
    If chk_basic_sunday.Value = 0 Then
        If DropAllComma(txt_main_salary_sunday_to.Text) <> DropAllComma(txt_main_salary_sunday.Text) Then
            SQL = "SELECT (SELECT employee_code FROM m_salary_standard " & _
                    "WHERE employee_code = a.employee_code " & _
                        "AND date(salary_date) <= '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date desc limit 1) " & _
                  "FROM m_salary_standard a " & _
                  "WHERE basic_salary_sunday = " & DropAllComma(txt_main_salary_sunday.Text) & " " & _
                  "GROUP BY employee_code"
                  
'            SQL = "SELECT  DISTINCT employee_code FROM m_salary_standard " & _
'                  "WHERE basic_salary_sunday = " & DropAllComma(txt_main_salary_sunday.Text) & ""
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = 0
                ProgressBar1.Max = rscari.RecordCount
                
                lbl_ket.Visible = True
                lbl_ket.Caption = "Mengupdate Gaji Pokok Hari Minggu..."
                
                vBasicSunday = "- Gaji Pokok Hari Minggu = " & rscari.RecordCount & " Karyawan"
                
                rscari.MoveFirst
                While Not rscari.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
                    SQL = "SELECT * FROM m_salary_standard " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        SQL = "UPDATE m_salary_standard " & _
                              "SET basic_salary_sunday = '" & DropAllComma(txt_main_salary_sunday_to.Text) & "' " & _
                              "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & rscari.Fields(0).Value & "'"
                        CnG.Execute SQL
                    Else
                        SQL = "INSERT INTO m_salary_standard " & _
                              "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                                "flag_basic, basic_salary, '" & DropAllComma(txt_main_salary_sunday_to.Text) & "'," & _
                                "flag_presence, presence_allowance," & _
                                "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                                "flag_performance, performance_allowance, flag_family, family_allowance," & _
                                "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                                "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                                "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                              "FROM m_salary_standard " & _
                              "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                                "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                        CnG.Execute SQL
                    End If
                    rs.Close
                    
                rscari.MoveNext
                Wend
            End If
            rscari.Close
            
            ProgressBar1.Visible = False
            lbl_ket.Visible = False
        End If
    Else
        SQL = "SELECT DISTINCT employee_code FROM m_salary_standard"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Max = rscari.RecordCount
            
            lbl_ket.Visible = True
            lbl_ket.Caption = "Mengupdate Gaji Pokok Hari Minggu..."
            
            vBasicSunday = "- Gaji Pokok Hari Minggu = " & rscari.RecordCount & " Karyawan"
                
            rscari.MoveFirst
            While Not rscari.EOF
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                SQL = "SELECT * FROM m_salary_standard " & _
                      "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                            "AND employee_code = '" & rscari.Fields(0).Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    SQL = "UPDATE m_salary_standard " & _
                          "SET basic_salary_sunday = '" & DropAllComma(txt_main_salary_sunday_to.Text) & "' " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_salary_standard " & _
                          "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                            "flag_basic, basic_salary, '" & DropAllComma(txt_main_salary_sunday_to.Text) & "'," & _
                            "flag_presence, presence_allowance," & _
                            "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                            "flag_performance, performance_allowance, flag_family, family_allowance," & _
                            "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                            "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                            "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                          "FROM m_salary_standard " & _
                          "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    CnG.Execute SQL
                End If
                rs.Close
                
            rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        ProgressBar1.Visible = False
        lbl_ket.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '++++++++++++++++++ Tunj. Kehadiran  +++++++++++++++++++++
    If chk_hadir.Value = 0 Then
        If DropAllComma(txt_presence_allowance_to.Text) <> DropAllComma(txt_presence_allowance.Text) Then
            SQL = "SELECT (SELECT employee_code FROM m_salary_standard " & _
                    "WHERE employee_code = a.employee_code " & _
                        "AND date(salary_date) <= '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date desc limit 1) " & _
                  "FROM m_salary_standard a " & _
                  "WHERE presence_allowance = " & DropAllComma(txt_presence_allowance.Text) & " " & _
                  "GROUP BY employee_code"
                  
'            SQL = "SELECT  DISTINCT employee_code FROM m_salary_standard " & _
'                  "WHERE presence_allowance = " & DropAllComma(txt_presence_allowance.Text) & ""
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = 0
                ProgressBar1.Max = rscari.RecordCount
                
                lbl_ket.Visible = True
                lbl_ket.Caption = "Mengupdate Tunjangan Kehadiran..."
                
                vHadir = "- Tunjangan Kehadiran = " & rscari.RecordCount & " Karyawan"
                
                rscari.MoveFirst
                While Not rscari.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
                    SQL = "SELECT * FROM m_salary_standard " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        SQL = "UPDATE m_salary_standard " & _
                              "SET presence_allowance = '" & DropAllComma(txt_presence_allowance_to.Text) & "' " & _
                              "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & rscari.Fields(0).Value & "'"
                        CnG.Execute SQL
                    Else
                        SQL = "INSERT INTO m_salary_standard " & _
                              "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                                "flag_basic, basic_salary,basic_salary_sunday," & _
                                "flag_presence, '" & DropAllComma(txt_presence_allowance_to.Text) & "'," & _
                                "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                                "flag_performance, performance_allowance, flag_family, family_allowance," & _
                                "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                                "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                                "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                              "FROM m_salary_standard " & _
                              "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                                "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                        CnG.Execute SQL
                    End If
                    rs.Close
                    
                rscari.MoveNext
                Wend
            End If
            rscari.Close
            
            ProgressBar1.Visible = False
            lbl_ket.Visible = False
        End If
    Else
        SQL = "SELECT DISTINCT employee_code FROM m_salary_standard"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Max = rscari.RecordCount
            
            lbl_ket.Visible = True
            lbl_ket.Caption = "Mengupdate Tunjangan Kehadiran..."
            
            vHadir = "- Tunjangan Kehadiran = " & rscari.RecordCount & " Karyawan"
                
            rscari.MoveFirst
            While Not rscari.EOF
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                SQL = "SELECT * FROM m_salary_standard " & _
                      "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                            "AND employee_code = '" & rscari.Fields(0).Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    SQL = "UPDATE m_salary_standard " & _
                          "SET presence_allowance = '" & DropAllComma(txt_presence_allowance_to.Text) & "' " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_salary_standard " & _
                          "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                            "flag_basic, basic_salary, basic_salary_sunday," & _
                            "flag_presence, '" & DropAllComma(txt_presence_allowance_to.Text) & "'," & _
                            "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                            "flag_performance, performance_allowance, flag_family, family_allowance," & _
                            "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                            "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                            "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                          "FROM m_salary_standard " & _
                          "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    CnG.Execute SQL
                End If
                rs.Close
                
            rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        ProgressBar1.Visible = False
        lbl_ket.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++ Tunj. Insentif +++++++++++++++++++++++
    If chk_insentif.Value = 0 Then
        If DropAllComma(txt_incentive_allowance_to.Text) <> DropAllComma(txt_incentive_allowance.Text) Then
            SQL = "SELECT (SELECT employee_code FROM m_salary_standard " & _
                    "WHERE employee_code = a.employee_code " & _
                        "AND date(salary_date) <= '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date desc limit 1) " & _
                  "FROM m_salary_standard a " & _
                  "WHERE incentive_allowance = " & DropAllComma(txt_incentive_allowance.Text) & " " & _
                  "GROUP BY employee_code"
                  
'            SQL = "SELECT  DISTINCT employee_code FROM m_salary_standard " & _
'                  "WHERE incentive_allowance = " & DropAllComma(txt_incentive_allowance.Text) & ""
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = 0
                ProgressBar1.Max = rscari.RecordCount
                
                lbl_ket.Visible = True
                lbl_ket.Caption = "Mengupdate Tunjangan Insentif..."
                
                vInsentif = "- Tunjangan Insentif = " & rscari.RecordCount & " Karyawan"
                
                rscari.MoveFirst
                While Not rscari.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    SQL = "SELECT * FROM m_salary_standard " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        SQL = "UPDATE m_salary_standard " & _
                              "SET incentive_allowance = '" & DropAllComma(txt_incentive_allowance_to.Text) & "' " & _
                              "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & rscari.Fields(0).Value & "'"
                        CnG.Execute SQL
                    Else
                        SQL = "INSERT INTO m_salary_standard " & _
                              "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                                "flag_basic, basic_salary,basic_salary_sunday," & _
                                "flag_presence, presence_allowance," & _
                                "flag_incentive, '" & DropAllComma(txt_incentive_allowance_to.Text) & "', flag_title, title_allowance," & _
                                "flag_performance, performance_allowance, flag_family, family_allowance," & _
                                "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                                "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                                "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                              "FROM m_salary_standard " & _
                              "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                                "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                        CnG.Execute SQL
                    End If
                    rs.Close
                    
                rscari.MoveNext
                Wend
            End If
            rscari.Close
            
            ProgressBar1.Visible = False
            lbl_ket.Visible = False
        End If
    Else
        SQL = "SELECT DISTINCT employee_code FROM m_salary_standard"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Max = rscari.RecordCount
            
            lbl_ket.Visible = True
            lbl_ket.Caption = "Mengupdate Tunjangan Insentif..."
            
            vInsentif = "- Tunjangan Insentif = " & rscari.RecordCount & " Karyawan"
                
            rscari.MoveFirst
            While Not rscari.EOF
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                SQL = "SELECT * FROM m_salary_standard " & _
                      "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                            "AND employee_code = '" & rscari.Fields(0).Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    SQL = "UPDATE m_salary_standard " & _
                          "SET incentive_allowance = '" & DropAllComma(txt_incentive_allowance_to.Text) & "' " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_salary_standard " & _
                          "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                            "flag_basic, basic_salary, basic_salary_sunday," & _
                            "flag_presence, presence_allowance," & _
                            "flag_incentive, '" & DropAllComma(txt_incentive_allowance_to.Text) & "', flag_title, title_allowance," & _
                            "flag_performance, performance_allowance, flag_family, family_allowance," & _
                            "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                            "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                            "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                          "FROM m_salary_standard " & _
                          "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    CnG.Execute SQL
                End If
                rs.Close
                
            rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        ProgressBar1.Visible = False
        lbl_ket.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++ Tunj. Jabatan +++++++++++++++++++++++
    If chk_jabatan.Value = 0 Then
        If DropAllComma(txt_title_allowance_to.Text) <> DropAllComma(txt_title_allowance.Text) Then
            SQL = "SELECT (SELECT employee_code FROM m_salary_standard " & _
                    "WHERE employee_code = a.employee_code " & _
                        "AND date(salary_date) <= '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date desc limit 1) " & _
                  "FROM m_salary_standard a " & _
                  "WHERE title_allowance = " & DropAllComma(txt_title_allowance.Text) & " " & _
                  "GROUP BY employee_code"
                  
'            SQL = "SELECT  DISTINCT employee_code FROM m_salary_standard " & _
'                  "WHERE title_allowance = " & DropAllComma(txt_title_allowance.Text) & ""
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = 0
                ProgressBar1.Max = rscari.RecordCount
                
                lbl_ket.Visible = True
                lbl_ket.Caption = "Mengupdate Tunjangan Jabatan..."
                
                vJabatan = "- Tunjangan Jabatan = " & rscari.RecordCount & " Karyawan"
                
                rscari.MoveFirst
                While Not rscari.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
                    SQL = "SELECT * FROM m_salary_standard " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        SQL = "UPDATE m_salary_standard " & _
                              "SET title_allowance = '" & DropAllComma(txt_title_allowance_to.Text) & "' " & _
                              "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & rscari.Fields(0).Value & "'"
                        CnG.Execute SQL
                    Else
                        SQL = "INSERT INTO m_salary_standard " & _
                              "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                                "flag_basic, basic_salary,basic_salary_sunday," & _
                                "flag_presence, presence_allowance," & _
                                "flag_incentive, incentive_allowance, flag_title, '" & DropAllComma(txt_title_allowance_to.Text) & "'," & _
                                "flag_performance, performance_allowance, flag_family, family_allowance," & _
                                "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                                "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                                "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                              "FROM m_salary_standard " & _
                              "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                                "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                        CnG.Execute SQL
                    End If
                    rs.Close
                    
                rscari.MoveNext
                Wend
            End If
            rscari.Close
            
            ProgressBar1.Visible = False
            lbl_ket.Visible = False
        End If
    Else
        SQL = "SELECT DISTINCT employee_code FROM m_salary_standard"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Max = rscari.RecordCount
            
            lbl_ket.Visible = True
            lbl_ket.Caption = "Mengupdate Tunjangan Jabatan..."
            
            vJabatan = "- Tunjangan Jabatan = " & rscari.RecordCount & " Karyawan"
                
            rscari.MoveFirst
            While Not rscari.EOF
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                SQL = "SELECT * FROM m_salary_standard " & _
                      "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                            "AND employee_code = '" & rscari.Fields(0).Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    SQL = "UPDATE m_salary_standard " & _
                          "SET title_allowance = '" & DropAllComma(txt_title_allowance_to.Text) & "' " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_salary_standard " & _
                          "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                            "flag_basic, basic_salary, basic_salary_sunday," & _
                            "flag_presence, presence_allowance," & _
                            "flag_incentive, incentive_allowance, flag_title, '" & DropAllComma(txt_title_allowance_to.Text) & "'," & _
                            "flag_performance, performance_allowance, flag_family, family_allowance," & _
                            "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                            "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                            "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                          "FROM m_salary_standard " & _
                          "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    CnG.Execute SQL
                End If
                rs.Close
                
            rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        ProgressBar1.Visible = False
        lbl_ket.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++ Tunj. Prestasi  +++++++++++++++++++++
    If chk_prestasi.Value = 0 Then
        If DropAllComma(txt_performance_allowance_to.Text) <> DropAllComma(txt_performance_allowance.Text) Then
            SQL = "SELECT (SELECT employee_code FROM m_salary_standard " & _
                    "WHERE employee_code = a.employee_code " & _
                        "AND date(salary_date) <= '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date desc limit 1) " & _
                  "FROM m_salary_standard a " & _
                  "WHERE performance_allowance = " & DropAllComma(txt_performance_allowance.Text) & " " & _
                  "GROUP BY employee_code"
                  
'            SQL = "SELECT  DISTINCT employee_code FROM m_salary_standard " & _
'                  "WHERE performance_allowance = " & DropAllComma(txt_performance_allowance.Text) & ""
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = 0
                ProgressBar1.Max = rscari.RecordCount
                
                lbl_ket.Visible = True
                lbl_ket.Caption = "Mengupdate Tunjangan Prestasi..."
                
                vPrestasi = "- Tunjangan Prestasi = " & rscari.RecordCount & " Karyawan"
                
                rscari.MoveFirst
                While Not rscari.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
                    SQL = "SELECT * FROM m_salary_standard " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        SQL = "UPDATE m_salary_standard " & _
                              "SET performance_allowance = '" & DropAllComma(txt_performance_allowance_to.Text) & "' " & _
                              "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & rscari.Fields(0).Value & "'"
                        CnG.Execute SQL
                    Else
                        SQL = "INSERT INTO m_salary_standard " & _
                              "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                                "flag_basic, basic_salary,basic_salary_sunday," & _
                                "flag_presence, presence_allowance," & _
                                "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                                "flag_performance, '" & DropAllComma(txt_performance_allowance_to.Text) & "', flag_family, family_allowance," & _
                                "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                                "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                                "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                              "FROM m_salary_standard " & _
                              "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                                "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                        CnG.Execute SQL
                    End If
                    rs.Close
                    
                rscari.MoveNext
                Wend
            End If
            rscari.Close
            
            ProgressBar1.Visible = False
            lbl_ket.Visible = False
        End If
    Else
        SQL = "SELECT DISTINCT employee_code FROM m_salary_standard"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Max = rscari.RecordCount
            
            lbl_ket.Visible = True
            lbl_ket.Caption = "Mengupdate Tunjangan Prestasi..."
            
            vPrestasi = "- Tunjangan Prestasi = " & rscari.RecordCount & " Karyawan"
                
            rscari.MoveFirst
            While Not rscari.EOF
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                SQL = "SELECT * FROM m_salary_standard " & _
                      "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                            "AND employee_code = '" & rscari.Fields(0).Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    SQL = "UPDATE m_salary_standard " & _
                          "SET performance_allowance = '" & DropAllComma(txt_performance_allowance_to.Text) & "' " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_salary_standard " & _
                          "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                            "flag_basic, basic_salary, basic_salary_sunday," & _
                            "flag_presence, presence_allowance," & _
                            "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                            "flag_performance, '" & DropAllComma(txt_performance_allowance_to.Text) & "', flag_family, family_allowance," & _
                            "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                            "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                            "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                          "FROM m_salary_standard " & _
                          "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    CnG.Execute SQL
                End If
                rs.Close
                
            rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        ProgressBar1.Visible = False
        lbl_ket.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++ Tunj. Keluarga  +++++++++++++++++++++
    If chk_keluarga.Value = 0 Then
        If DropAllComma(txt_family_allowance_to.Text) <> DropAllComma(txt_family_allowance.Text) Then
            SQL = "SELECT (SELECT employee_code FROM m_salary_standard " & _
                    "WHERE employee_code = a.employee_code " & _
                        "AND date(salary_date) <= '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date desc limit 1) " & _
                  "FROM m_salary_standard a " & _
                  "WHERE family_allowance = " & DropAllComma(txt_family_allowance.Text) & " " & _
                  "GROUP BY employee_code"
                  
'            SQL = "SELECT  DISTINCT employee_code FROM m_salary_standard " & _
'                  "WHERE family_allowance = " & DropAllComma(txt_family_allowance.Text) & ""
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = 0
                ProgressBar1.Max = rscari.RecordCount
                
                lbl_ket.Visible = True
                lbl_ket.Caption = "Mengupdate Tunjangan Keluarga..."
                
                vKeluarga = "- Tunjangan Keluarga = " & rscari.RecordCount & " Karyawan"
                
                rscari.MoveFirst
                While Not rscari.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
                    SQL = "SELECT * FROM m_salary_standard " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        SQL = "UPDATE m_salary_standard " & _
                              "SET family_allowance = '" & DropAllComma(txt_family_allowance_to.Text) & "' " & _
                              "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & rscari.Fields(0).Value & "'"
                        CnG.Execute SQL
                    Else
                        SQL = "INSERT INTO m_salary_standard " & _
                              "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                                "flag_basic, basic_salary,basic_salary_sunday," & _
                                "flag_presence, presence_allowance," & _
                                "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                                "flag_performance, performance_allowance, flag_family, '" & DropAllComma(txt_family_allowance_to.Text) & "'," & _
                                "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                                "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                                "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                              "FROM m_salary_standard " & _
                              "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                                "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                        CnG.Execute SQL
                    End If
                    rs.Close
                    
                rscari.MoveNext
                Wend
            End If
            rscari.Close
            
            ProgressBar1.Visible = False
            lbl_ket.Visible = False
        End If
    Else
        SQL = "SELECT DISTINCT employee_code FROM m_salary_standard"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Max = rscari.RecordCount
            
            lbl_ket.Visible = True
            lbl_ket.Caption = "Mengupdate Tunjangan Keluarga..."
            
            vKeluarga = "- Tunjangan Keluarga = " & rscari.RecordCount & " Karyawan"
            
            rscari.MoveFirst
            While Not rscari.EOF
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                SQL = "SELECT * FROM m_salary_standard " & _
                      "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                            "AND employee_code = '" & rscari.Fields(0).Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    SQL = "UPDATE m_salary_standard " & _
                          "SET family_allowance = '" & DropAllComma(txt_family_allowance_to.Text) & "' " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_salary_standard " & _
                          "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                            "flag_basic, basic_salary, basic_salary_sunday," & _
                            "flag_presence, presence_allowance," & _
                            "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                            "flag_performance, performance_allowance, flag_family, '" & DropAllComma(txt_family_allowance_to.Text) & "'," & _
                            "flag_meal, meal_allowance, shift2_allowance, shift3_allowance,pph21_allowance," & _
                            "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                            "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                          "FROM m_salary_standard " & _
                          "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    CnG.Execute SQL
                End If
                rs.Close
                
            rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        ProgressBar1.Visible = False
        lbl_ket.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '++++++++++++++++++ Tunj. Trans/Makan  +++++++++++++++++++
    If chk_trans_makan.Value = 0 Then
        If DropAllComma(txt_meal_allowance_to.Text) <> DropAllComma(txt_meal_allowance.Text) Then
            SQL = "SELECT (SELECT employee_code FROM m_salary_standard " & _
                    "WHERE employee_code = a.employee_code " & _
                        "AND date(salary_date) <= '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date desc limit 1) " & _
                  "FROM m_salary_standard a " & _
                  "WHERE meal_allowance = " & DropAllComma(txt_meal_allowance.Text) & " " & _
                  "GROUP BY employee_code"
                  
'            SQL = "SELECT  DISTINCT employee_code FROM m_salary_standard " & _
'                  "WHERE meal_allowance = " & DropAllComma(txt_meal_allowance.Text) & ""
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = 0
                ProgressBar1.Max = rscari.RecordCount
                
                lbl_ket.Visible = True
                lbl_ket.Caption = "Mengupdate Tunjangan Transportasi/Makan..."
                
                vTrans = "- Tunjangan Transportasi/Makan = " & rscari.RecordCount & " Karyawan"
                
                rscari.MoveFirst
                While Not rscari.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
                    SQL = "SELECT * FROM m_salary_standard " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        SQL = "UPDATE m_salary_standard " & _
                              "SET meal_allowance = '" & DropAllComma(txt_meal_allowance_to.Text) & "' " & _
                              "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & rscari.Fields(0).Value & "'"
                        CnG.Execute SQL
                    Else
                        SQL = "INSERT INTO m_salary_standard " & _
                              "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                                "flag_basic, basic_salary,basic_salary_sunday," & _
                                "flag_presence, presence_allowance," & _
                                "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                                "flag_performance, performance_allowance, flag_family, family_allowance," & _
                                "flag_meal, '" & DropAllComma(txt_meal_allowance_to.Text) & "', shift2_allowance, shift3_allowance,pph21_allowance," & _
                                "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                                "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                              "FROM m_salary_standard " & _
                              "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                                "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                        CnG.Execute SQL
                    End If
                    rs.Close
                    
                rscari.MoveNext
                Wend
            End If
            rscari.Close
            
            ProgressBar1.Visible = False
            lbl_ket.Visible = False
        End If
    Else
        SQL = "SELECT DISTINCT employee_code FROM m_salary_standard"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Max = rscari.RecordCount
            
            lbl_ket.Visible = True
            lbl_ket.Caption = "Mengupdate Tunjangan Transportasi/Makan..."
            
            vTrans = "- Tunjangan Transportasi/Makan = " & rscari.RecordCount & " Karyawan"
                
            rscari.MoveFirst
            While Not rscari.EOF
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                SQL = "SELECT * FROM m_salary_standard " & _
                      "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                            "AND employee_code = '" & rscari.Fields(0).Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    SQL = "UPDATE m_salary_standard " & _
                          "SET meal_allowance = '" & DropAllComma(txt_meal_allowance_to.Text) & "' " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_salary_standard " & _
                          "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                            "flag_basic, basic_salary, basic_salary_sunday," & _
                            "flag_presence, presence_allowance," & _
                            "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                            "flag_performance, performance_allowance, flag_family, family_allowance," & _
                            "flag_meal, '" & DropAllComma(txt_meal_allowance_to.Text) & "', shift2_allowance, shift3_allowance,pph21_allowance," & _
                            "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                            "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                          "FROM m_salary_standard " & _
                          "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    CnG.Execute SQL
                End If
                rs.Close
                
            rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        ProgressBar1.Visible = False
        lbl_ket.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '++++++++++++++++++++ Tunj. Shift 2  +++++++++++++++++++++
    If chk_shift2.Value = 0 Then
        If DropAllComma(txt_shift2_allowance_to.Text) <> DropAllComma(txt_shift2_allowance.Text) Then
            SQL = "SELECT (SELECT employee_code FROM m_salary_standard " & _
                    "WHERE employee_code = a.employee_code " & _
                        "AND date(salary_date) <= '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date desc limit 1) " & _
                  "FROM m_salary_standard a " & _
                  "WHERE shift2_allowance = " & DropAllComma(txt_shift2_allowance.Text) & " " & _
                  "GROUP BY employee_code"
                  
'            SQL = "SELECT  DISTINCT employee_code FROM m_salary_standard " & _
'                  "WHERE shift2_allowance = " & DropAllComma(txt_shift2_allowance.Text) & ""
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = 0
                ProgressBar1.Max = rscari.RecordCount
                
                lbl_ket.Visible = True
                lbl_ket.Caption = "Mengupdate Tunjangan Shift 2..."
                
                vShift2 = "- Tunjangan Shift 2 = " & rscari.RecordCount & " Karyawan"
                
                rscari.MoveFirst
                While Not rscari.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
                    SQL = "SELECT * FROM m_salary_standard " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        SQL = "UPDATE m_salary_standard " & _
                              "SET shift2_allowance = '" & DropAllComma(txt_shift2_allowance_to.Text) & "' " & _
                              "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & rscari.Fields(0).Value & "'"
                        CnG.Execute SQL
                    Else
                        SQL = "INSERT INTO m_salary_standard " & _
                              "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                                "flag_basic, basic_salary,basic_salary_sunday," & _
                                "flag_presence, presence_allowance," & _
                                "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                                "flag_performance, performance_allowance, flag_family, family_allowance," & _
                                "flag_meal, meal_allowance, '" & DropAllComma(txt_shift2_allowance_to.Text) & "'," & _
                                "shift3_allowance,pph21_allowance," & _
                                "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                                "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                              "FROM m_salary_standard " & _
                              "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                                "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                        CnG.Execute SQL
                    End If
                    rs.Close
                    
                rscari.MoveNext
                Wend
            End If
            rscari.Close
            
            ProgressBar1.Visible = False
            lbl_ket.Visible = False
        End If
    Else
        SQL = "SELECT DISTINCT employee_code FROM m_salary_standard"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Max = rscari.RecordCount
            
            lbl_ket.Visible = True
            lbl_ket.Caption = "Mengupdate Tunjangan Shift 2..."
            
            vShift2 = "- Tunjangan Shift 2 = " & rscari.RecordCount & " Karyawan"
                
            rscari.MoveFirst
            While Not rscari.EOF
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                SQL = "SELECT * FROM m_salary_standard " & _
                      "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                            "AND employee_code = '" & rscari.Fields(0).Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    SQL = "UPDATE m_salary_standard " & _
                          "SET shift2_allowance = '" & DropAllComma(txt_shift2_allowance_to.Text) & "' " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_salary_standard " & _
                          "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                            "flag_basic, basic_salary, basic_salary_sunday," & _
                            "flag_presence, presence_allowance," & _
                            "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                            "flag_performance, performance_allowance, flag_family, family_allowance," & _
                            "flag_meal, meal_allowance, '" & DropAllComma(txt_shift2_allowance_to.Text) & "'," & _
                            "shift3_allowance,pph21_allowance," & _
                            "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                            "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                          "FROM m_salary_standard " & _
                          "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    CnG.Execute SQL
                End If
                rs.Close
                
            rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        ProgressBar1.Visible = False
        lbl_ket.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '++++++++++++++++++++ Tunj. Shift 3  +++++++++++++++++++++
    If chk_shift3.Value = 0 Then
        If DropAllComma(txt_shift3_allowance_to.Text) <> DropAllComma(txt_shift3_allowance.Text) Then
            SQL = "SELECT (SELECT employee_code FROM m_salary_standard " & _
                    "WHERE employee_code = a.employee_code " & _
                        "AND date(salary_date) <= '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date desc limit 1) " & _
                  "FROM m_salary_standard a " & _
                  "WHERE shift3_allowance = " & DropAllComma(txt_shift3_allowance.Text) & " " & _
                  "GROUP BY employee_code"
                  
'            SQL = "SELECT  DISTINCT employee_code FROM m_salary_standard " & _
'                  "WHERE shift3_allowance = " & DropAllComma(txt_shift3_allowance.Text) & ""
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = 0
                ProgressBar1.Max = rscari.RecordCount
                
                lbl_ket.Visible = True
                lbl_ket.Caption = "Mengupdate Tunjangan Shift 3..."
                
                vShift3 = "- Tunjangan Shift 3 = " & rscari.RecordCount & " Karyawan"
                
                rscari.MoveFirst
                While Not rscari.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
                    SQL = "SELECT * FROM m_salary_standard " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        SQL = "UPDATE m_salary_standard " & _
                              "SET shift3_allowance = '" & DropAllComma(txt_shift3_allowance_to.Text) & "' " & _
                              "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & rscari.Fields(0).Value & "'"
                        CnG.Execute SQL
                    Else
                        SQL = "INSERT INTO m_salary_standard " & _
                              "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                                "flag_basic, basic_salary,basic_salary_sunday," & _
                                "flag_presence, presence_allowance," & _
                                "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                                "flag_performance, performance_allowance, flag_family, family_allowance," & _
                                "flag_meal, meal_allowance, shift2_allowance," & _
                                "'" & DropAllComma(txt_shift3_allowance_to.Text) & "',pph21_allowance," & _
                                "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                                "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                              "FROM m_salary_standard " & _
                              "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                                "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                        CnG.Execute SQL
                    End If
                    rs.Close
                    
                rscari.MoveNext
                Wend
            End If
            rscari.Close
            
            ProgressBar1.Visible = False
            lbl_ket.Visible = False
        End If
    Else
        SQL = "SELECT DISTINCT employee_code FROM m_salary_standard"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Max = rscari.RecordCount
            
            lbl_ket.Visible = True
            lbl_ket.Caption = "Mengupdate Tunjangan Shift 3..."
            
            vShift3 = "- Tunjangan Shift 3 = " & rscari.RecordCount & " Karyawan"
                
            rscari.MoveFirst
            While Not rscari.EOF
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                SQL = "SELECT * FROM m_salary_standard " & _
                      "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                            "AND employee_code = '" & rscari.Fields(0).Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    SQL = "UPDATE m_salary_standard " & _
                          "SET shift3_allowance = '" & DropAllComma(txt_shift3_allowance_to.Text) & "' " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_salary_standard " & _
                          "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                            "flag_basic, basic_salary, basic_salary_sunday," & _
                            "flag_presence, presence_allowance," & _
                            "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                            "flag_performance, performance_allowance, flag_family, family_allowance," & _
                            "flag_meal, meal_allowance, shift2_allowance," & _
                            "'" & DropAllComma(txt_shift3_allowance_to.Text) & "',pph21_allowance," & _
                            "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                            "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                          "FROM m_salary_standard " & _
                          "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    CnG.Execute SQL
                End If
                rs.Close
                
            rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        ProgressBar1.Visible = False
        lbl_ket.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++++ Tunj. PPh21  ++++++++++++++++++++++
    If chk_pph21.Value = 0 Then
        If DropAllComma(txt_pph21_allowance_to.Text) <> DropAllComma(txt_pph21_allowance.Text) Then
            SQL = "SELECT (SELECT employee_code FROM m_salary_standard " & _
                    "WHERE employee_code = a.employee_code " & _
                        "AND date(salary_date) <= '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date desc limit 1) " & _
                  "FROM m_salary_standard a " & _
                  "WHERE pph21_allowance = " & DropAllComma(txt_pph21_allowance.Text) & " " & _
                  "GROUP BY employee_code"
                  
'            SQL = "SELECT  DISTINCT employee_code FROM m_salary_standard " & _
'                  "WHERE pph21_allowance = " & DropAllComma(txt_pph21_allowance.Text) & ""
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = 0
                ProgressBar1.Max = rscari.RecordCount
                
                lbl_ket.Visible = True
                lbl_ket.Caption = "Mengupdate Tunjangan PPh21..."
                
                vPPh21 = "- Tunjangan PPh21 = " & rscari.RecordCount & " Karyawan"
                
                rscari.MoveFirst
                While Not rscari.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
                    SQL = "SELECT * FROM m_salary_standard " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        SQL = "UPDATE m_salary_standard " & _
                              "SET pph21_allowance = '" & DropAllComma(txt_pph21_allowance_to.Text) & "' " & _
                              "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & rscari.Fields(0).Value & "'"
                        CnG.Execute SQL
                    Else
                        SQL = "INSERT INTO m_salary_standard " & _
                              "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                                "flag_basic, basic_salary,basic_salary_sunday," & _
                                "flag_presence, presence_allowance," & _
                                "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                                "flag_performance, performance_allowance, flag_family, family_allowance," & _
                                "flag_meal, meal_allowance, shift2_allowance," & _
                                "shift3_allowance,'" & DropAllComma(txt_pph21_allowance_to.Text) & "'," & _
                                "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                                "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                              "FROM m_salary_standard " & _
                              "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                                "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                        CnG.Execute SQL
                    End If
                    rs.Close
                    
                rscari.MoveNext
                Wend
            End If
            rscari.Close
            
            ProgressBar1.Visible = False
            lbl_ket.Visible = False
        End If
    Else
        SQL = "SELECT DISTINCT employee_code FROM m_salary_standard"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Max = rscari.RecordCount
            
            lbl_ket.Visible = True
            lbl_ket.Caption = "Mengupdate Tunjangan PPh21..."
            
            vPPh21 = "- Tunjangan PPh21 = " & rscari.RecordCount & " Karyawan"
                
            rscari.MoveFirst
            While Not rscari.EOF
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                SQL = "SELECT * FROM m_salary_standard " & _
                      "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                            "AND employee_code = '" & rscari.Fields(0).Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    SQL = "UPDATE m_salary_standard " & _
                          "SET pph21_allowance = '" & DropAllComma(txt_pph21_allowance_to.Text) & "' " & _
                          "WHERE DATE(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
                                "AND employee_code = '" & rscari.Fields(0).Value & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_salary_standard " & _
                          "SELECT '" & rscari.Fields(0).Value & "', '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
                            "flag_basic, basic_salary, basic_salary_sunday," & _
                            "flag_presence, presence_allowance," & _
                            "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
                            "flag_performance, performance_allowance, flag_family, family_allowance," & _
                            "flag_meal, meal_allowance, shift2_allowance," & _
                            "shift3_allowance,'" & DropAllComma(txt_pph21_allowance_to.Text) & "'," & _
                            "late_time_tolerance, late_amount, pph21_type, ptkp_type, jstk_type, flag_gaji," & _
                            "Now(),'" & LOGIN_CODE & "',NULL,NULL " & _
                          "FROM m_salary_standard " & _
                          "WHERE employee_code = '" & rscari.Fields(0).Value & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    CnG.Execute SQL
                End If
                rs.Close
                
            rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        ProgressBar1.Visible = False
        lbl_ket.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    CnG.CommitTrans
    Me.MousePointer = vbNormal
    
    MsgBox "Update Data Berhasil Dengan Rincian : " & _
            IIf(vBasic <> "", Chr(13) & vBasic, "") & _
            IIf(vBasicSunday <> "", Chr(13) & vBasicSunday, "") & _
            IIf(vHadir <> "", Chr(13) & vHadir, "") & _
            IIf(vInsentif <> "", Chr(13) & vInsentif, "") & _
            IIf(vJabatan <> "", Chr(13) & vJabatan, "") & _
            IIf(vPrestasi <> "", Chr(13) & vPrestasi, "") & _
            IIf(vKeluarga <> "", Chr(13) & vKeluarga, "") & _
            IIf(vTrans <> "", Chr(13) & vTrans, "") & _
            IIf(vShift2 <> "", Chr(13) & vShift2, "") & _
            IIf(vShift3 <> "", Chr(13) & vShift3, "") & _
            IIf(vPPh21 <> "", Chr(13) & vPPh21, ""), vbInformation, headerMSG
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    cboLeave.ListIndex = 0
    cboSalary.ListIndex = 0
    
    Call load_data
    Call load_data_company
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
End Sub

Private Sub clear_filter()
    If SSTab1.Tab = 0 Then
        For Each Col In TDBGrid2.Columns
            Col.FilterText = ""
        Next Col
        rsJHK.Filter = adFilterNone
    ElseIf SSTab1.Tab = 1 Then
        For Each Col In TDBGrid1.Columns
            Col.FilterText = ""
        Next Col
        rsUMK.Filter = adFilterNone
    ElseIf SSTab1.Tab = 2 Then
        For Each Col In TDBGrid3.Columns
            Col.FilterText = ""
        Next Col
        rsLeave.Filter = adFilterNone
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

Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_title = Nothing
End Sub

Private Sub filter_change()
On Error GoTo Err

    Dim i As Integer
    
    If SSTab1.Tab = 0 Then
        Set Cols = TDBGrid2.Columns
        i = TDBGrid2.Col
        TDBGrid2.HoldFields
        
        rsJHK.Filter = getFilter()
        TDBGrid2.Col = i
        TDBGrid2.EditActive = True
        
        TDBGrid2.SelStart = Len(TDBGrid2.Columns(i).FilterText)
        If TDBGrid2.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid2.Col = i
        End If
    ElseIf SSTab1.Tab = 1 Then
        Set Cols = TDBGrid1.Columns
        i = TDBGrid1.Col
        TDBGrid1.HoldFields
        
        rsUMK.Filter = getFilter()
        TDBGrid1.Col = i
        TDBGrid1.EditActive = True
        
        TDBGrid1.SelStart = Len(TDBGrid1.Columns(i).FilterText)
        If TDBGrid1.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid1.Col = i
        End If
    ElseIf SSTab1.Tab = 2 Then
        Set Cols = TDBGrid3.Columns
        i = TDBGrid3.Col
        TDBGrid3.HoldFields
        
        rsLeave.Filter = getFilter()
        TDBGrid3.Col = i
        TDBGrid3.EditActive = True
        
        TDBGrid3.SelStart = Len(TDBGrid3.Columns(i).FilterText)
        If TDBGrid3.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid3.Col = i
        End If
    End If

    Exit Sub

Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data()
    If SSTab1.Tab = 0 Then
        If rsJHK.State Then rsJHK.Close
        SQL = "select a.*,b.company_name,c.division_name from m_pref_jhk a join m_company b on a.company_code = b.company_code " & _
                "LEFT JOIN m_division c on a.company_code = c.company_code and a.division_code = c.division_code"
        rsJHK.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid2.DataSource = rsJHK
    ElseIf SSTab1.Tab = 1 Then
        If rsUMK.State Then rsUMK.Close
        SQL = "select * from m_pref_umk "
        rsUMK.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid1.DataSource = rsUMK
    ElseIf SSTab1.Tab = 2 Then
        If rsLeave.State Then rsLeave.Close
        SQL = "select a.*,CASE WHEN leave_type = 0 then 'PERIODE KERJA' else 'TAHUN' end type,b.company_name,c.division_name " & _
                "from m_pref_leave a join m_company b on a.company_code = b.company_code " & _
                "LEFT JOIN m_division c on a.company_code = c.company_code and a.division_code = c.division_code"
        rsLeave.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid3.DataSource = rsLeave
    ElseIf SSTab1.Tab = 3 Then
        If rsGeneral.State Then rsGeneral.Close
        SQL = "select * from m_pref_gen"
        rsGeneral.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        Call set_edit_data
    ElseIf SSTab1.Tab = 4 Then
        If rsPrefTunj.State Then rsPrefTunj.Close
        SQL = "select * from m_pref_tunj"
        rsPrefTunj.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        Call set_edit_data
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call load_data
    
    If SSTab1.Tab = 0 Or SSTab1.Tab = 2 Then
        Call load_data_company
        
        int_mode = 0
        Call load_mode
    ElseIf SSTab1.Tab = 1 Then
        int_mode = 0
        Call load_mode
    End If
End Sub

Private Sub TDBCombo_company_jhk_ItemChange()
    If TDBCombo_company_jhk.ApproxCount > 0 Then
        TDBCombo_company_jhk.Text = TDBCombo_company_jhk.Columns("company_code").Value
        txt_company_name_jhk.Text = TDBCombo_company_jhk.Columns("company_name").Value

        Call load_data_division
    End If
End Sub

Private Sub TDBCombo_company_leave_ItemChange()
    If TDBCombo_company_leave.ApproxCount > 0 Then
        TDBCombo_company_leave.Text = TDBCombo_company_leave.Columns("company_code").Value
        txt_company_name_leave.Text = TDBCombo_company_leave.Columns("company_name").Value

        Call load_data_division
    End If
End Sub

Private Sub tdbcombo_division_jhk_itemChange()
    If TDBCombo_division_jhk.ApproxCount > 0 Then
        TDBCombo_division_jhk.Text = TDBCombo_division_jhk.Columns("division_code").Value
        txt_division_name_jhk.Text = TDBCombo_division_jhk.Columns("division_name").Value
    End If
End Sub

Private Sub tdbcombo_division_leave_itemChange()
    If TDBCombo_division_leave.ApproxCount > 0 Then
        TDBCombo_division_leave.Text = TDBCombo_division_leave.Columns("division_code").Value
        txt_division_name_leave.Text = TDBCombo_division_leave.Columns("division_name").Value
    End If
End Sub

Public Sub load_data_company()
    If SSTab1.Tab = 0 Then
        TDBCombo_company_jhk.Text = "": txt_company_name_jhk = ""
        
        If rsCompany.State Then rsCompany.Close
        SQL = "select * from m_company order by company_code"
        rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBCombo_company_jhk.RowSource = rsCompany
    ElseIf SSTab1.Tab = 2 Then
        TDBCombo_company_leave.Text = "": txt_company_name_leave = ""
        
        If rsCompany.State Then rsCompany.Close
        SQL = "select * from m_company order by company_code"
        rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBCombo_company_leave.RowSource = rsCompany
    End If
End Sub

Public Sub load_data_division()
    If SSTab1.Tab = 0 Then
        TDBCombo_division_jhk.Text = "": txt_division_name_jhk.Text = ""
        
        If rsDivision.State Then rsDivision.Close
        SQL = "select * from m_division where company_code = '" & TDBCombo_company_jhk.Text & "' order by company_code"
        rsDivision.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBCombo_division_jhk.RowSource = rsDivision
    ElseIf SSTab1.Tab = 2 Then
        TDBCombo_division_leave.Text = "": txt_division_name_leave.Text = ""
        
        If rsDivision.State Then rsDivision.Close
        SQL = "select * from m_division where company_code = '" & TDBCombo_company_leave.Text & "' order by company_code"
        rsDivision.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBCombo_division_leave.RowSource = rsDivision
    End If
End Sub

Private Sub txt_umk_value_Validate(Cancel As Boolean)
    If Not Trim(txt_umk_value) = "" Then
        txt_umk_value = FormatNumber(DropAllComma(txt_umk_value))
    End If
End Sub


Private Sub TDBGrid2_FilterChange()
    Call filter_change
End Sub

Private Sub TDBGrid1_FilterChange()
    Call filter_change
End Sub

Private Sub TDBGrid3_FilterChange()
    Call filter_change
End Sub

Private Sub cmdNew_JHK_Click()
    Call new_data
End Sub

Private Sub cmdSave_JHK_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_JHK_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_JHK_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_JHK_Click()
    Call cancel_data
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

Private Sub cmdNew_Leave_Click()
    Call new_data
End Sub

Private Sub cmdSave_Leave_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Leave_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Leave_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Leave_Click()
    Call cancel_data
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

Private Sub txt_incentive_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_presence_allowance) = "" Then
        txt_incentive_allowance = FormatNumber(DropAllComma(txt_incentive_allowance))
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


Private Sub txt_main_salary_to_Validate(Cancel As Boolean)
    If Not Trim(txt_main_salary_to) = "" Then
        txt_main_salary_to = FormatNumber(DropAllComma(txt_main_salary_to))
    End If
End Sub

Private Sub txt_main_salary_sunday_to_Validate(Cancel As Boolean)
    If Not Trim(txt_main_salary_sunday_to) = "" Then
        txt_main_salary_sunday_to = FormatNumber(DropAllComma(txt_main_salary_sunday_to))
    End If
End Sub

Private Sub txt_presence_allowance_to_Validate(Cancel As Boolean)
    If Not Trim(txt_presence_allowance_to) = "" Then
        txt_presence_allowance_to = FormatNumber(DropAllComma(txt_presence_allowance_to))
    End If
End Sub

Private Sub txt_incentive_allowance_to_Validate(Cancel As Boolean)
    If Not Trim(txt_incentive_allowance_to) = "" Then
        txt_incentive_allowance_to = FormatNumber(DropAllComma(txt_incentive_allowance_to))
    End If
End Sub

Private Sub txt_title_allowance_to_Validate(Cancel As Boolean)
    If Not Trim(txt_title_allowance_to) = "" Then
        txt_title_allowance_to = FormatNumber(DropAllComma(txt_title_allowance_to))
    End If
End Sub

Private Sub txt_performance_allowance_to_Validate(Cancel As Boolean)
    If Not Trim(txt_performance_allowance_to) = "" Then
        txt_performance_allowance_to = FormatNumber(DropAllComma(txt_performance_allowance_to))
    End If
End Sub

Private Sub txt_family_allowance_to_Validate(Cancel As Boolean)
    If Not Trim(txt_family_allowance_to) = "" Then
        txt_family_allowance_to = FormatNumber(DropAllComma(txt_family_allowance_to))
    End If
End Sub

Private Sub txt_meal_allowance_to_Validate(Cancel As Boolean)
    If Not Trim(txt_meal_allowance_to) = "" Then
        txt_meal_allowance_to = FormatNumber(DropAllComma(txt_meal_allowance_to))
    End If
End Sub

Private Sub txt_shift2_allowance_to_Validate(Cancel As Boolean)
    If Not Trim(txt_shift2_allowance_to) = "" Then
        txt_shift2_allowance_to = FormatNumber(DropAllComma(txt_shift2_allowance_to))
    End If
End Sub

Private Sub txt_shift3_allowance_to_Validate(Cancel As Boolean)
    If Not Trim(txt_shift3_allowance_to) = "" Then
        txt_shift3_allowance_to = FormatNumber(DropAllComma(txt_shift3_allowance_to))
    End If
End Sub

Private Sub txt_pph21_allowance_to_Validate(Cancel As Boolean)
    If Not Trim(txt_pph21_allowance) = "" Then
        txt_pph21_allowance_to = FormatNumber(DropAllComma(txt_pph21_allowance_to))
    End If
End Sub
