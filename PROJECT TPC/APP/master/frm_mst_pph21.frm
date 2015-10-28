VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_pph21 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER TAX"
   ClientHeight    =   8955
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_mst_pph21.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "frm_mst_pph21.frx":058A
   ScaleHeight     =   8955
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   90
      TabIndex        =   1
      Top             =   870
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "GENERAL"
      TabPicture(0)   =   "frm_mst_pph21.frx":0B14
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame20"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "P T K P"
      TabPicture(1)   =   "frm_mst_pph21.frx":0B30
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label27"
      Tab(1).Control(1)=   "TDBGrid_PTKP"
      Tab(1).Control(2)=   "fra_entry_dtl_ptkp"
      Tab(1).Control(3)=   "TDBCombo_ptkp"
      Tab(1).Control(4)=   "txt_ptkp"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(6)=   "fra_entry_ptkp"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "PPh 21"
      TabPicture(2)   =   "frm_mst_pph21.frx":0B4C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label49"
      Tab(2).Control(1)=   "TDBGrid_PPh"
      Tab(2).Control(2)=   "fra_entry_dtl_pph"
      Tab(2).Control(3)=   "TDBCombo_pph"
      Tab(2).Control(4)=   "txt_pph21"
      Tab(2).Control(5)=   "frmTombol"
      Tab(2).Control(6)=   "fra_entry_pph"
      Tab(2).ControlCount=   7
      Begin VB.Frame Frame20 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   1740
         TabIndex        =   118
         Top             =   4380
         Width           =   7965
         Begin prj_tpc.vbButton cmdSave 
            Height          =   705
            Left            =   6000
            TabIndex        =   119
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":0B68
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":0B94
            picn            =   "frm_mst_pph21.frx":0BB2
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
      Begin VB.Frame Frame21 
         Height          =   3255
         Left            =   1740
         TabIndex        =   114
         Top             =   1020
         Width           =   7965
         Begin VB.OptionButton optYear 
            Caption         =   "2014"
            Height          =   255
            Index           =   1
            Left            =   4890
            TabIndex        =   115
            Top             =   1260
            Width           =   1035
         End
         Begin VB.OptionButton optYear 
            Caption         =   "2008"
            Height          =   255
            Index           =   0
            Left            =   4020
            TabIndex        =   116
            Top             =   1260
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.Label Label63 
            Caption         =   "PPh 21 REPORT METHOD"
            Height          =   285
            Left            =   1830
            TabIndex        =   117
            Top             =   1260
            Width           =   2025
         End
      End
      Begin VB.Frame fra_entry_ptkp 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   96
         Top             =   3510
         Visible         =   0   'False
         Width           =   11295
         Begin VB.TextBox txt_ptkp_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   10
            TabIndex        =   99
            Top             =   600
            Width           =   1695
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
            TabIndex        =   98
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_name_ptkp 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   97
            Top             =   960
            Width           =   3495
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "PTKP NAME"
            Height          =   195
            Left            =   3480
            TabIndex        =   101
            Top             =   960
            Width           =   825
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "PTKP CODE"
            Height          =   195
            Left            =   3480
            TabIndex        =   100
            Top             =   600
            Width           =   825
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   88
         Top             =   5460
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew_PTKP 
            Height          =   705
            Left            =   540
            TabIndex        =   89
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&New Dtl"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":1C46
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":1C72
            picn            =   "frm_mst_pph21.frx":1C90
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_PTKP 
            Height          =   705
            Left            =   1560
            TabIndex        =   90
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save Dtl"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":2D24
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":2D50
            picn            =   "frm_mst_pph21.frx":2D6E
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_PTKP 
            Height          =   705
            Left            =   2580
            TabIndex        =   91
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Edit Dtl"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":3E02
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":3E2E
            picn            =   "frm_mst_pph21.frx":3E4C
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_PTKP 
            Height          =   705
            Left            =   3600
            TabIndex        =   92
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Delete Dtl"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":4EE0
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":4F0C
            picn            =   "frm_mst_pph21.frx":4F2A
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_PTKP 
            Height          =   705
            Left            =   4620
            TabIndex        =   93
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Cancel Dtl"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":5FBE
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":5FEA
            picn            =   "frm_mst_pph21.frx":6008
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton CmdNew_Master_PTKP 
            Height          =   705
            Left            =   8850
            TabIndex        =   94
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&New"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":709C
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":70C8
            picn            =   "frm_mst_pph21.frx":70E6
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_All_PTKP 
            Height          =   705
            Left            =   9870
            TabIndex        =   95
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Delete"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":817A
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":81A6
            picn            =   "frm_mst_pph21.frx":81C4
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
      Begin VB.TextBox txt_ptkp 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   -71970
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   86
         Top             =   600
         Width           =   3855
      End
      Begin VB.Frame fra_entry_pph 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   66
         Top             =   3540
         Width           =   11295
         Begin VB.TextBox txt_pph_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   10
            TabIndex        =   68
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txt_pph_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   67
            Top             =   960
            Width           =   3495
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "PPh NAME"
            Height          =   195
            Left            =   3480
            TabIndex        =   70
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "PPh CODE"
            Height          =   195
            Left            =   3480
            TabIndex        =   69
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   58
         Top             =   5460
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew_Pph 
            Height          =   705
            Left            =   540
            TabIndex        =   59
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&New Dtl"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":9258
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":9284
            picn            =   "frm_mst_pph21.frx":92A2
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_Pph 
            Height          =   705
            Left            =   1560
            TabIndex        =   60
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save Dtl"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":A336
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":A362
            picn            =   "frm_mst_pph21.frx":A380
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_Pph 
            Height          =   705
            Left            =   2580
            TabIndex        =   61
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Edit Dtl"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":B414
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":B440
            picn            =   "frm_mst_pph21.frx":B45E
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_Pph 
            Height          =   705
            Left            =   3600
            TabIndex        =   62
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Delete Dtl"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":C4F2
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":C51E
            picn            =   "frm_mst_pph21.frx":C53C
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_Pph 
            Height          =   705
            Left            =   4620
            TabIndex        =   63
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Cancel Dtl"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":D5D0
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":D5FC
            picn            =   "frm_mst_pph21.frx":D61A
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton CmdNew_Master_Pph 
            Height          =   705
            Left            =   8850
            TabIndex        =   64
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&New"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":E6AE
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":E6DA
            picn            =   "frm_mst_pph21.frx":E6F8
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_All_Pph 
            Height          =   705
            Left            =   9870
            TabIndex        =   65
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Delete"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":F78C
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":F7B8
            picn            =   "frm_mst_pph21.frx":F7D6
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
      Begin VB.TextBox txt_pph21 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   -71970
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   56
         Top             =   600
         Width           =   3855
      End
      Begin VB.Frame fra_entry_Department 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   47
         Top             =   2280
         Width           =   11205
         Begin VB.TextBox txt_department_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   50
            TabIndex        =   50
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txt_department_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   50
            TabIndex        =   49
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txt_department_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   10
            TabIndex        =   48
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   3480
            TabIndex        =   53
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   3480
            TabIndex        =   52
            Top             =   600
            Width           =   510
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   3480
            TabIndex        =   51
            Top             =   960
            Width           =   525
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   38
         Top             =   4590
         Width           =   11205
         Begin VB.Timer Timer7 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_tpc.vbButton cmdNew_Department 
            Height          =   705
            Left            =   540
            TabIndex        =   39
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&New"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":1086A
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":10896
            picn            =   "frm_mst_pph21.frx":108B4
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_Department 
            Height          =   705
            Left            =   1560
            TabIndex        =   40
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":11948
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":11974
            picn            =   "frm_mst_pph21.frx":11992
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_Department 
            Height          =   705
            Left            =   2580
            TabIndex        =   41
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Edit"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":12A26
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":12A52
            picn            =   "frm_mst_pph21.frx":12A70
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_Department 
            Height          =   705
            Left            =   3600
            TabIndex        =   42
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Delete"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":13B04
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":13B30
            picn            =   "frm_mst_pph21.frx":13B4E
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_Department 
            Height          =   705
            Left            =   4620
            TabIndex        =   43
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Cancel"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":14BE2
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":14C0E
            picn            =   "frm_mst_pph21.frx":14C2C
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
      Begin VB.Frame fra_entry_Division 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   28
         Top             =   1770
         Width           =   11175
         Begin VB.TextBox txt_department 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5400
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   32
            Top             =   660
            Width           =   3855
         End
         Begin VB.TextBox txt_division_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   31
            Top             =   1770
            Width           =   3495
         End
         Begin VB.TextBox txt_division_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   30
            Top             =   1410
            Width           =   3495
         End
         Begin VB.TextBox txt_division_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   29
            Top             =   1050
            Width           =   1695
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_department 
            Height          =   375
            Left            =   3600
            OleObjectBlob   =   "frm_mst_pph21.frx":15CC0
            TabIndex        =   33
            Top             =   660
            Width           =   1725
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "DEPARTMENT*"
            Height          =   195
            Left            =   2280
            TabIndex        =   37
            Top             =   690
            Width           =   1185
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   2280
            TabIndex        =   36
            Top             =   1770
            Width           =   1095
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   2280
            TabIndex        =   35
            Top             =   1050
            Width           =   510
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   2280
            TabIndex        =   34
            Top             =   1410
            Width           =   525
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   22
         Top             =   4560
         Width           =   11175
         Begin VB.Timer Timer3 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_tpc.vbButton cmdNew_Division 
            Height          =   705
            Left            =   540
            TabIndex        =   23
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&New"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":17C29
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":17C55
            picn            =   "frm_mst_pph21.frx":17C73
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_Division 
            Height          =   705
            Left            =   1560
            TabIndex        =   24
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":18D07
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":18D33
            picn            =   "frm_mst_pph21.frx":18D51
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_Division 
            Height          =   705
            Left            =   2580
            TabIndex        =   25
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Edit"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":19DE5
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":19E11
            picn            =   "frm_mst_pph21.frx":19E2F
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_Division 
            Height          =   705
            Left            =   3600
            TabIndex        =   26
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Delete"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":1AEC3
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":1AEEF
            picn            =   "frm_mst_pph21.frx":1AF0D
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_Division 
            Height          =   705
            Left            =   4620
            TabIndex        =   27
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Cancel"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":1BFA1
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":1BFCD
            picn            =   "frm_mst_pph21.frx":1BFEB
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
      Begin VB.Frame fra_entry_Grade 
         Height          =   2655
         Left            =   -74790
         TabIndex        =   15
         Top             =   1770
         Width           =   11175
         Begin VB.TextBox txt_grade_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   10
            TabIndex        =   18
            Top             =   810
            Width           =   1695
         End
         Begin VB.TextBox txt_grade_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   50
            TabIndex        =   17
            Top             =   1170
            Width           =   3495
         End
         Begin VB.TextBox txt_grade_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   50
            TabIndex        =   16
            Top             =   1530
            Width           =   3495
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   2910
            TabIndex        =   21
            Top             =   1170
            Width           =   525
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   2910
            TabIndex        =   20
            Top             =   810
            Width           =   510
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   2910
            TabIndex        =   19
            Top             =   1530
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74790
         TabIndex        =   9
         Top             =   4560
         Width           =   11175
         Begin VB.Timer Timer4 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_tpc.vbButton cmdNew_Grade 
            Height          =   705
            Left            =   540
            TabIndex        =   10
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&New"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":1D07F
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":1D0AB
            picn            =   "frm_mst_pph21.frx":1D0C9
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_Grade 
            Height          =   705
            Left            =   1560
            TabIndex        =   11
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":1E15D
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":1E189
            picn            =   "frm_mst_pph21.frx":1E1A7
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_Grade 
            Height          =   705
            Left            =   2580
            TabIndex        =   12
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Edit"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":1F23B
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":1F267
            picn            =   "frm_mst_pph21.frx":1F285
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_Grade 
            Height          =   705
            Left            =   3600
            TabIndex        =   13
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Delete"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":20319
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":20345
            picn            =   "frm_mst_pph21.frx":20363
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_Grade 
            Height          =   705
            Left            =   4620
            TabIndex        =   14
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Cancel"
            enab            =   -1  'True
            font            =   "frm_mst_pph21.frx":213F7
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_pph21.frx":21423
            picn            =   "frm_mst_pph21.frx":21441
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
      Begin VB.Frame fra_entry_level 
         Height          =   2775
         Left            =   -74730
         TabIndex        =   2
         Top             =   1650
         Width           =   11175
         Begin VB.TextBox txt_level_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1620
            Width           =   5145
         End
         Begin VB.TextBox txt_level_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1230
            Width           =   3795
         End
         Begin VB.TextBox txt_level_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   10
            TabIndex        =   3
            Top             =   810
            Width           =   1215
         End
         Begin VB.Label tlabel26 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   1950
            TabIndex        =   8
            Top             =   1650
            Width           =   1095
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   1950
            TabIndex        =   7
            Top             =   840
            Width           =   510
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   1950
            TabIndex        =   6
            Top             =   1260
            Width           =   525
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
         Height          =   3915
         Left            =   -74760
         TabIndex        =   44
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
         TabIndex        =   45
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
         TabIndex        =   46
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
         TabIndex        =   54
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
         Left            =   -73770
         OleObjectBlob   =   "frm_mst_pph21.frx":224D5
         TabIndex        =   57
         Top             =   600
         Width           =   1695
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_ptkp 
         Height          =   375
         Left            =   -73770
         OleObjectBlob   =   "frm_mst_pph21.frx":2442F
         TabIndex        =   87
         Top             =   600
         Width           =   1695
      End
      Begin VB.Frame fra_entry_dtl_pph 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   71
         Top             =   2580
         Width           =   11295
         Begin VB.CheckBox chk_flag_over 
            Height          =   255
            Left            =   6120
            TabIndex        =   77
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txt_pph21_percentage 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   76
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txt_pph21_upper 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            TabIndex        =   75
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txt_pph_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   74
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txt_pph21_under 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            TabIndex        =   73
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txt_pph21_number 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   72
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "PERCENTAGE"
            Height          =   195
            Left            =   4800
            TabIndex        =   83
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "UP"
            Height          =   195
            Left            =   4800
            TabIndex        =   82
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "TO VALUE"
            Height          =   195
            Left            =   4800
            TabIndex        =   81
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   4800
            TabIndex        =   80
            Top             =   1920
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "NO."
            Height          =   195
            Left            =   600
            TabIndex        =   79
            Top             =   480
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "FROM VALUE"
            Height          =   195
            Left            =   4800
            TabIndex        =   78
            Top             =   480
            Width           =   945
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_PPh 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   84
         Top             =   1020
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
      Begin VB.Frame fra_entry_dtl_ptkp 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   102
         Top             =   2940
         Width           =   11295
         Begin VB.TextBox txt_ptkp_number 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   10
            TabIndex        =   107
            Top             =   600
            Width           =   1695
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
            TabIndex        =   106
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_ptkp_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   105
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txt_ptkp_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   104
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txt_ptkp_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            TabIndex        =   103
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            Height          =   195
            Left            =   3480
            TabIndex        =   111
            Top             =   960
            Width           =   420
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "NO."
            Height          =   195
            Left            =   3480
            TabIndex        =   110
            Top             =   600
            Width           =   285
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   3480
            TabIndex        =   109
            Top             =   1680
            Width           =   1020
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "VALUE"
            Height          =   195
            Left            =   3480
            TabIndex        =   108
            Top             =   1320
            Width           =   465
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_PTKP 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   112
         Top             =   1020
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7646
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NUMBER"
         Columns(0).DataField=   "ptkp_number"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "PTKP NAME"
         Columns(1).DataField=   "ptkp_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "VALUE"
         Columns(2).DataField=   "ptkp_value"
         Columns(2).NumberFormat=   "Standard"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2275"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2196"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=5980"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5900"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=4868"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4789"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=5768"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=5689"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
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
         Caption         =   "LIST OF PTKP"
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
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
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
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "PTKP TYPE"
         Height          =   195
         Left            =   -74760
         TabIndex        =   113
         Top             =   630
         Width           =   765
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "PPH21 TYPE"
         Height          =   195
         Left            =   -74760
         TabIndex        =   85
         Top             =   630
         Width           =   870
      End
   End
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   10680
      TabIndex        =   55
      Top             =   8100
      Width           =   945
      _extentx        =   1667
      _extenty        =   1244
      btype           =   14
      tx              =   "&Exit"
      enab            =   -1  'True
      font            =   "frm_mst_pph21.frx":2638A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   15790320
      bcolo           =   15790320
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frm_mst_pph21.frx":263B6
      picn            =   "frm_mst_pph21.frx":263D4
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER TAX"
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
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_pph21.frx":27468
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11970
   End
End
Attribute VB_Name = "frm_mst_pph21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPTKP As New ADODB.Recordset
Dim rsPTKP_Detail As New ADODB.Recordset
Dim rspph As New ADODB.Recordset
Dim rsPPh_Detail As New ADODB.Recordset

Dim rsPPh_Gen As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim v_value_under, v_value_upper, v_value_percent, v_value As Double

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
    check_validate_exist_new = False
    
    If SSTab1.Tab = 1 Then
        SQL = "select count(ptkp_number) as rec_count from m_ptkp_detail where ptkp_number = " _
                & "'" & Int(txt_ptkp_number) & "' AND ptkp_code = '" & TDBCombo_ptkp.Text & "'"
    ElseIf SSTab1.Tab = 2 Then
        SQL = "select count(pph21_number) as rec_count from m_pph21_detail where pph21_number = " _
                & "'" & Int(txt_pph21_number) & "' AND pph21_code = '" & TDBCombo_pph.Text & "'"
    End If
    
    rs.Open SQL, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        Exit Function
    End If
End Function

Private Sub check_invalid()
    MsgBox "Data found!", vbCritical, headerMSG
    If SSTab1.Tab = 1 Then
        txt_ptkp_number = ""
        If txt_ptkp_number.Enabled = True Then txt_ptkp_number.SetFocus
    ElseIf SSTab1.Tab = 2 Then
        txt_pph21_number = ""
        If txt_pph21_number.Enabled = True Then txt_pph21_number.SetFocus
    End If
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
        
    If SSTab1.Tab = 1 Then
        If Not txt_ptkp_number = rsPTKP_Detail.Fields("ptkp_number").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        If Not txt_pph21_number = rsPPh_Detail.Fields("pph21_number").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    End If
    
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True

    If SSTab1.Tab = 1 Then
        If Trim(txt_ptkp_number) = "" Then
            MsgBox "Number is empty!", vbOKOnly + vbInformation, headerMSG
            txt_ptkp_number.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If Trim(txt_ptkp_name) = "" Then
            MsgBox "Name is empty!", vbOKOnly + vbInformation, headerMSG
            txt_ptkp_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        If Trim(txt_pph21_number) = "" Then
            MsgBox "Number is empty!", vbOKOnly + vbInformation, headerMSG
            txt_pph21_number.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If Trim(txt_pph21_under) = "" Then
            MsgBox "Start from value is empty!", vbOKOnly + vbInformation, headerMSG
            txt_pph21_under.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If Trim(txt_pph21_upper) = "" Then
            MsgBox "End value is empty!", vbOKOnly + vbInformation, headerMSG
            txt_pph21_upper.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If Trim(txt_pph21_percentage) = "" Then
            MsgBox "Percentage value is empty!", vbOKOnly + vbInformation, headerMSG
            txt_pph21_percentage.SetFocus
            check_validate_new = False
            Exit Function
        End If
    End If
End Function

Private Sub cancel_data()
    int_mode = 0
    Call load_mode
    CmdNew_Master_Pph.Caption = "&New"
    CmdNew_Master_PTKP.Caption = "&New"
    cmdCancel_Pph.Caption = "&Cancel Dtl"
    cmdCancel_PTKP.Caption = "&Cancel Dtl"
End Sub

Private Sub delete_all_data()
Dim i As Integer

On Error GoTo Err
    If SSTab1.Tab = 1 Then
        i = MsgBox("Are you sure want to delete data '" _
            & txt_ptkp.Text & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_ptkp_detail where " _
                & "ptkp_code = '" & TDBCombo_ptkp.Text & "'"
        CnG.Execute "delete from m_ptkp where ptkp_code = " _
                & "'" & TDBCombo_ptkp.Text & "'"
        
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        SQL = "Update temp_sal_proses set salary_proses = 0"
        CnG.Execute SQL
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        CnG.CommitTrans
        
        Call load_data_master
        Call load_data_detail
        int_mode = 0
        Call load_mode
        
        TDBCombo_ptkp.Text = ""
        txt_name_ptkp = ""
        Set TDBGrid_PTKP.DataSource = Nothing
    ElseIf SSTab1.Tab = 2 Then
        i = MsgBox("Are you sure want to delete data '" _
            & txt_pph21.Text & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_pph21_detail where " _
                & "pph21_code = '" & TDBCombo_pph.Text & "'"
        CnG.Execute "delete from m_pph21 where pph21_code = " _
                & "'" & TDBCombo_pph.Text & "'"
        
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        SQL = "Update temp_sal_proses set salary_proses = 0"
        CnG.Execute SQL
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        CnG.CommitTrans
        
        Call load_data_master
        Call load_data_detail
        int_mode = 0
        Call load_mode
        
        TDBCombo_pph.Text = ""
        txt_pph21 = ""
        Set TDBGrid_PPh.DataSource = Nothing
    End If
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub delete_data()
Dim i As Integer
On Error GoTo Err
    If SSTab1.Tab = 1 Then
        If Not (TDBGrid_PTKP.ApproxCount > 0 And TDBGrid_PTKP.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Are you sure want to delete data '" _
            & TDBGrid_PTKP.Columns("ptkp_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_ptkp_detail where ptkp_number = " _
                & "'" & TDBGrid_PTKP.Columns("ptkp_number").Value & "' AND ptkp_code = '" & TDBCombo_ptkp.Text & "'"
        
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        SQL = "Update temp_sal_proses set salary_proses = 0"
        CnG.Execute SQL
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        CnG.CommitTrans
        
        Call load_data_detail
        int_mode = 0
        Call load_mode
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_PPh.ApproxCount > 0 And TDBGrid_PPh.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Are you sure want to delete data '" _
            & TDBGrid_PPh.Columns("pph21_number").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_pph21_detail where pph21_number = " _
                & "'" & TDBGrid_PPh.Columns("pph21_number").Value & "' AND pph21_code = '" & TDBCombo_pph.Text & "'"
        
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        SQL = "Update temp_sal_proses set salary_proses = 0"
        CnG.Execute SQL
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        CnG.CommitTrans
        
        Call load_data_detail
        int_mode = 0
        Call load_mode
    End If
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub set_edit_data()
    vSetData = 1
    If SSTab1.Tab = 0 Then
        Dim vYear As Integer
        If rsPPh_Gen.RecordCount > 0 Then
            With rsPPh_Gen
                vYear = .Fields("int_report_method").Value
                optYear(0).Value = IIf(vYear = 0, True, False)
                optYear(1).Value = IIf(vYear = 0, False, True)
            End With
        End If
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_PTKP.ApproxCount > 0 And TDBGrid_PTKP.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsPTKP_Detail
            txt_ptkp_number = .Fields("ptkp_number").Value
            txt_ptkp_name = .Fields("ptkp_name").Value
            txt_ptkp_value = FormatNumber(.Fields("ptkp_value").Value)
            txt_ptkp_description = .Fields("description").Value
        End With
        
            v_value = txt_ptkp_value
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_PPh.ApproxCount > 0 And TDBGrid_PPh.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsPPh_Detail
            txt_pph21_number = .Fields("pph21_number").Value
            txt_pph21_under = FormatNumber(.Fields("pph21_under").Value)
            txt_pph21_upper = FormatNumber(.Fields("pph21_upper").Value)
            chk_flag_over.Value = .Fields("flag_over").Value
            txt_pph21_percentage = FormatNumber(.Fields("pph21_percentage").Value)
            txt_pph_description = .Fields("description").Value
        End With
        
            v_value_under = txt_pph21_under
            v_value_upper = txt_pph21_upper
            v_value_percent = txt_pph21_percentage
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

Private Sub new_data_master()
On Error GoTo Err
    If SSTab1.Tab = 1 Then
        If CmdNew_Master_PTKP.Caption = "&New" Then
            CmdNew_Master_PTKP.Caption = "&Save"
            cmdCancel_PTKP.Caption = "&Cancel"
            Call set_buttons_enable(False, False, False, False, True, False, False)
            fra_entry_ptkp.Visible = True
            fra_entry_dtl_ptkp.Visible = False
            
            txt_ptkp_code.Text = ""
            txt_name_ptkp.Text = ""
            txt_ptkp_code.SetFocus
            
            CmdNew_Master_PTKP.Enabled = True
            
        Else
            SQL = "INSERT INTO m_ptkp(ptkp_code,ptkp_name) " _
                    & "VALUES ('" & txt_ptkp_code & "','" & txt_name_ptkp & "')"
            CnG.Execute SQL
            
            Call set_buttons_enable(True, False, True, True, False, True, True)
            CmdNew_Master_PTKP.Caption = "&New"
            cmdCancel_PTKP.Caption = "&Cancel Dtl"
                
            fra_entry_ptkp.Visible = False
            
            txt_ptkp_code.Text = ""
            txt_name_ptkp.Text = ""
        
            Call load_data_master
            txt_ptkp.Text = ""
        End If
    ElseIf SSTab1.Tab = 2 Then
        If CmdNew_Master_Pph.Caption = "&New" Then
            CmdNew_Master_Pph.Caption = "&Save"
            cmdCancel_Pph.Caption = "&Cancel"
            Call set_buttons_enable(False, False, False, False, True, False, False)
            fra_entry_pph.Visible = True
            fra_entry_dtl_pph.Visible = False
            
            txt_pph_code.Text = ""
            txt_pph_name.Text = ""
            txt_pph_code.SetFocus
            
            CmdNew_Master_Pph.Enabled = True
            
        Else
            SQL = "INSERT INTO m_pph21(pph21_code,pph21_name) " _
                    & "VALUES ('" & txt_pph_code & "','" & txt_pph_name & "')"
            CnG.Execute SQL
            
            Call set_buttons_enable(True, False, True, True, False, True, True)
            CmdNew_Master_Pph.Caption = "&New"
            cmdCancel_Pph.Caption = "&Cancel Dtl"
                
            fra_entry_pph.Visible = False
            
            txt_pph_code.Text = ""
            txt_pph_name.Text = ""
        
            Call load_data_master
            txt_pph21.Text = ""
        End If
    End If
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub insert_new_data()
On Error GoTo Err
    CnG.BeginTrans
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    SQL = "Update temp_sal_proses set salary_proses = 0"
    CnG.Execute SQL
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    If SSTab1.Tab = 1 Then
        SQL = "INSERT INTO m_ptkp_detail (ptkp_code,ptkp_number,ptkp_name," & _
                "ptkp_value,description) " & _
              "VALUES( " & _
                "'" & Trim(TDBCombo_ptkp.Text) & "','" & Int(txt_ptkp_number.Text) & "'," & _
                "'" & Trim(txt_ptkp_name) & "','" & Val(DropAllComma(txt_ptkp_value)) & "'," & _
                "'" & Trim(txt_ptkp_description) & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "INSERT INTO m_pph21_detail (pph21_code,pph21_number,pph21_under," & _
                "pph21_upper,flag_over,pph21_percentage,description) " & _
              "VALUES( " & _
                "'" & Trim(TDBCombo_pph.Text) & "','" & Trim(txt_pph21_number.Text) & "'," & _
                "'" & Val(DropAllComma(txt_pph21_under)) & "','" & Val(DropAllComma(txt_pph21_upper)) & "'," & _
                "'" & IIf(chk_flag_over, 1, 0) & "','" & Val(DropAllComma(txt_pph21_percentage)) & "'," & _
                "'" & Trim(txt_pph_description) & "')"
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
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    If SSTab1.Tab = 1 Then
        If v_value <> txt_ptkp_value.Text Then
            SQL = "Update temp_sal_proses set salary_proses = 0"
            CnG.Execute SQL
        End If
        
        SQL = "UPDATE m_ptkp_detail SET ptkp_code = '" & TDBCombo_ptkp.Text & "'," & _
                "ptkp_number = '" & Int(txt_ptkp_number.Text) & "'," & _
                "ptkp_name = '" & Trim(txt_ptkp_name.Text) & "'," & _
                "ptkp_value = '" & Val(DropAllComma(txt_ptkp_value.Text)) & "'," & _
                "description = '" & Trim(txt_ptkp_description.Text) & "' " & _
              "WHERE ptkp_code = '" & TDBCombo_ptkp.Text & "' " & _
                "AND ptkp_number = '" & Int(txt_ptkp_number.Text) & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        If v_value_under <> txt_pph21_under.Text Or v_value_upper <> txt_pph21_upper.Text Or v_value_percent <> txt_pph21_percentage.Text Then
            SQL = "Update temp_sal_proses set salary_proses = 0"
            CnG.Execute SQL
        End If
        
        SQL = "UPDATE m_pph21_detail SET pph21_code = '" & Trim(TDBCombo_pph.Text) & "'," & _
                "pph21_number = '" & Trim(txt_pph21_number.Text) & "'," & _
                "pph21_under = '" & Val(DropAllComma(txt_pph21_under.Text)) & "'," & _
                "pph21_upper = '" & Val(DropAllComma(txt_pph21_upper.Text)) & "'," & _
                "flag_over = '" & IIf(chk_flag_over, 1, 0) & "'," & _
                "pph21_percentage = '" & Val(DropAllComma(txt_pph21_percentage.Text)) & "'," & _
                "description = '" & Trim(txt_pph_description.Text) & "' " & _
              "WHERE pph21_code = '" & Trim(TDBCombo_pph.Text) & "' " & _
                "AND pph21_number = '" & Trim(txt_pph21_number.Text) & "'"
        CnG.Execute SQL
    End If
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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

    Call load_data_detail
    int_mode = 0
    Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal F As Boolean, ByVal g As Boolean)
    If SSTab1.Tab = 1 Then
        cmdNew_PTKP.Enabled = a And blnUser_Add
        cmdSave_PTKP.Enabled = b
        cmdEdit_PTKP.Enabled = c And blnUser_Edit
        cmdDelete_PTKP.Enabled = d And blnUser_Delete
        cmdCancel_PTKP.Enabled = e
        
        CmdNew_Master_PTKP.Enabled = a And blnUser_Add
        cmdDelete_All_PTKP.Enabled = d And blnUser_Delete
    ElseIf SSTab1.Tab = 2 Then
        cmdNew_Pph.Enabled = a And blnUser_Add
        cmdSave_Pph.Enabled = b
        cmdEdit_Pph.Enabled = c And blnUser_Edit
        cmdDelete_Pph.Enabled = d And blnUser_Delete
        cmdCancel_Pph.Enabled = e
        
        CmdNew_Master_Pph.Enabled = a And blnUser_Add
        cmdDelete_All_Pph.Enabled = d And blnUser_Delete
    End If
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If SSTab1.Tab = 1 Then
                If Not LCase(Ctr.name) = "txt_ptkp" Then Ctr.Text = ""
            ElseIf SSTab1.Tab = 2 Then
                If Not LCase(Ctr.name) = "txt_pph21" Then Ctr.Text = ""
            End If
        ElseIf TypeOf Ctr Is TDBCombo Then
            If SSTab1.Tab = 1 Then
                If Not LCase(Ctr.name) = "tdbcombo_ptkp" Then Ctr.Text = ""
            ElseIf SSTab1.Tab = 2 Then
                If Not LCase(Ctr.name) = "tdbcombo_pph" Then Ctr.Text = ""
            End If
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

End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW Rule
    Call clear_view_data
    If SSTab1.Tab = 1 Then
        If Trim(TDBCombo_ptkp) = "" Then
            MsgBox "PTKP Type is not selected!", vbOKOnly + vbInformation, headerMSG
            TDBCombo_ptkp.SetFocus
            
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
    
        fra_entry_dtl_ptkp.Visible = True
        cmdCancel_PTKP.Caption = "&Cancel Dtl"
        txt_ptkp_number.Enabled = True
        TDBGrid_PTKP.Enabled = False
        Call set_new_data
        
        If txt_ptkp_number.Enabled = True Then
            txt_ptkp_number.SetFocus
        End If
    ElseIf SSTab1.Tab = 2 Then
        If Trim(TDBCombo_pph) = "" Then
            MsgBox "PPh Type is not selected!", vbOKOnly + vbInformation, headerMSG
            TDBCombo_pph.SetFocus
            
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
    
        fra_entry_dtl_pph.Visible = True
        cmdCancel_Pph.Caption = "&Cancel Dtl"
        txt_pph21_number.Enabled = True
        TDBGrid_PPh.Enabled = False
        Call set_new_data
        
        If txt_pph21_number.Enabled = True Then
            txt_pph21_number.SetFocus
        End If
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    If SSTab1.Tab = 1 Then
        fra_entry_ptkp.Visible = False
        fra_entry_dtl_ptkp.Visible = False
        TDBGrid_PTKP.Enabled = True
    ElseIf SSTab1.Tab = 2 Then
        fra_entry_pph.Visible = False
        fra_entry_dtl_pph.Visible = False
        TDBGrid_PPh.Enabled = True
    End If

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    
    If vSetData = 0 Then
        int_mode = 0
        Call load_mode
        Exit Sub
    End If
    
    If SSTab1.Tab = 1 Then
        txt_ptkp_number.Enabled = False
        fra_entry_dtl_ptkp.Visible = True
        TDBGrid_PTKP.Enabled = False
    ElseIf SSTab1.Tab = 2 Then
        txt_pph21_number.Enabled = False
        fra_entry_dtl_pph.Visible = True
        TDBGrid_PPh.Enabled = False
    End If
End If
End Sub

Private Sub load_mode()
    If int_mode = 1 Then        ' new Rule
        Call set_buttons_enable(False, True, False, False, True, False, False)
    ElseIf int_mode = 0 Then    ' view
        Call set_buttons_enable(True, False, True, True, False, True, True)
    ElseIf int_mode = 2 Then    ' edit/revise
        Call set_buttons_enable(False, True, False, False, True, False, False)
    End If
    
    Call set_data_mode
End Sub

Private Sub Form_Load()
    
    Call load_data_master
    oClause = ""
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    
    SSTab1.Tab = 0
    
    If SSTab1.Tab = 0 Then
        Call load_data_detail
    End If
End Sub

Private Sub clear_filter()
    If SSTab1.Tab = 1 Then
        For Each Col In TDBGrid_PTKP.Columns
            Col.FilterText = ""
        Next Col
        rsPTKP_Detail.Filter = adFilterNone
    ElseIf SSTab1.Tab = 2 Then
        For Each Col In TDBGrid_PPh.Columns
            Col.FilterText = ""
        Next Col
        rsPPh_Detail.Filter = adFilterNone
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
On Error GoTo Err

Dim i As Integer
    If SSTab1.Tab = 1 Then
        Set Cols = TDBGrid_PTKP.Columns
        i = TDBGrid_PTKP.Col
        TDBGrid_PTKP.HoldFields
        
        rsPTKP_Detail.Filter = getFilter()
        TDBGrid_PTKP.Col = i
        TDBGrid_PTKP.EditActive = True
        
        TDBGrid_PTKP.SelStart = Len(TDBGrid_PTKP.Columns(i).FilterText)
        If TDBGrid_PTKP.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_PTKP.Col = i
        End If
    ElseIf SSTab1.Tab = 2 Then
        Set Cols = TDBGrid_PPh.Columns
        i = TDBGrid_PPh.Col
        TDBGrid_PPh.HoldFields
        
        rsPPh_Detail.Filter = getFilter()
        TDBGrid_PPh.Col = i
        TDBGrid_PPh.EditActive = True
        
        TDBGrid_PPh.SelStart = Len(TDBGrid_PPh.Columns(i).FilterText)
        If TDBGrid_PPh.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_PPh.Col = i
        End If
    End If
    
    Exit Sub
    
Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data_detail()
    If SSTab1.Tab = 0 Then
        If rsPPh_Gen.State Then rsPPh_Gen.Close
        SQL = "SELECT int_report_method FROM m_pph21_gen"
        rsPPh_Gen.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        Call set_edit_data
    ElseIf SSTab1.Tab = 1 Then
        If rsPTKP_Detail.State Then rsPTKP_Detail.Close
        SQL = "select * from m_ptkp_detail " & _
                "where ptkp_code = '" & TDBCombo_ptkp.Text & "' " & oClause
        rsPTKP_Detail.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_PTKP.DataSource = rsPTKP_Detail
    ElseIf SSTab1.Tab = 2 Then
        If rsPPh_Detail.State Then rsPPh_Detail.Close
        SQL = "select * from m_pph21_detail where pph21_code = '" _
                & TDBCombo_pph.Columns("pph21_code").Value & "' " & oClause
        rsPPh_Detail.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_PPh.DataSource = rsPPh_Detail
    End If
End Sub

Private Sub load_data_master()
    If SSTab1.Tab = 1 Then
        If rsPTKP.State Then rsPTKP.Close
        SQL = "select * from m_ptkp order by ptkp_code"
        rsPTKP.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBCombo_ptkp.RowSource = rsPTKP
    ElseIf SSTab1.Tab = 2 Then
        If rspph.State Then rspph.Close
        SQL = "select * from m_pph21 order by pph21_code"
        rspph.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBCombo_pph.RowSource = rspph
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_pph21 = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    oClause = ""
        
    If SSTab1.Tab = 0 Then
        Call load_data_detail
    ElseIf SSTab1.Tab = 1 Then
        int_mode = 0
        Call load_mode
        Call load_data_master
    
        TDBGrid_PPh.DataSource = Nothing
    Else
        int_mode = 0
        Call load_mode
        Call load_data_master
    
        TDBGrid_PTKP.DataSource = Nothing
    End If
End Sub

Private Sub TDBCombo_pph_ItemChange()
    If TDBCombo_pph.ApproxCount > 0 Then
        TDBCombo_pph.Text = TDBCombo_pph.Columns("pph21_code").Value
        txt_pph21 = TDBCombo_pph.Columns("pph21_name").Value
        
        Call load_data_detail
    End If
End Sub

Private Sub TDBCombo_ptkp_ItemChange()
    If TDBCombo_ptkp.ApproxCount > 0 Then
        TDBCombo_ptkp.Text = TDBCombo_ptkp.Columns("ptkp_code").Value
        txt_ptkp = TDBCombo_ptkp.Columns("ptkp_name").Value
        
        Call load_data_detail
    End If
End Sub

Private Sub CmdSave_Click()
    SQL = "UPDATE m_pph21_gen SET int_report_method = '" & IIf(optYear(0).Value, 0, 1) & "'"
    CnG.Execute SQL
    
    MsgBox "Save successfully...", vbInformation, headerMSG
End Sub


Private Sub cmdNew_PTKP_Click()
    Call new_data
End Sub

Private Sub cmdSave_PTKP_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_PTKP_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_PTKP_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_PTKP_Click()
    Call cancel_data
End Sub

Private Sub CmdNew_Master_PTKP_Click()
    Call new_data_master
End Sub

Private Sub cmdDelete_All_PTKP_Click()
    Call delete_all_data
End Sub


Private Sub cmdNew_PPh_Click()
    Call new_data
End Sub

Private Sub cmdSave_PPh_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_PPh_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_PPh_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_PPh_Click()
    Call cancel_data
End Sub

Private Sub CmdNew_Master_PPh_Click()
    Call new_data_master
End Sub

Private Sub cmdDelete_All_PPh_Click()
    Call delete_all_data
End Sub


Private Sub TDBGrid_PPh_FilterChange()
    Call filter_change
End Sub

Private Sub TDBGrid_PTKP_FilterChange()
    Call filter_change
End Sub


Private Sub txt_pph_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_pph_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_ptkp_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_ptkp_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_ptkp_value_Validate(Cancel As Boolean)
    If Not Trim(txt_ptkp_value.Text) = "" Then
        txt_ptkp_value.Text = FormatNumber(DropAllComma(txt_ptkp_value.Text))
    End If
End Sub

Private Sub txt_pph21_under_Validate(Cancel As Boolean)
    If Not Trim(txt_pph21_under.Text) = "" Then
        txt_pph21_under.Text = FormatNumber(DropAllComma(txt_pph21_under.Text))
    End If
End Sub

Private Sub txt_pph21_upper_Validate(Cancel As Boolean)
Dim vPPhFrom As Long
Dim vPPhTo As Long
    
    vPPhFrom = Val(Replace(txt_pph21_under.Text, ",", ""))
    vPPhTo = Val(Replace(txt_pph21_upper.Text, ",", ""))
    If vPPhFrom > vPPhTo And chk_flag_over = 0 Then
        MsgBox "To Value is smaller than From Value!", vbExclamation, headerMSG
        txt_pph21_upper.SetFocus
        Exit Sub
    End If
    
    If Not Trim(txt_pph21_upper.Text) = "" Then
        txt_pph21_upper.Text = FormatNumber(DropAllComma(txt_pph21_upper.Text))
    End If
End Sub

Private Sub TDBGrid_PTKP_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_PTKP.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_PTKP.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_PTKP.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_PTKP.Columns(ColIndex).DataField
    Call load_data_detail

End Sub


Private Sub TDBGrid_PPh_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_PPh.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_PPh.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_PPh.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_PPh.Columns(ColIndex).DataField
    Call load_data_detail

End Sub


