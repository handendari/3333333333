VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_bpjs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BPJS KESEHATAN"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14145
   Icon            =   "frm_trans_bpjs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   14145
   Begin prj_genting.vbButton cmdExit 
      Height          =   705
      Left            =   12990
      TabIndex        =   19
      Top             =   7890
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
      MICON           =   "frm_trans_bpjs.frx":058A
      PICN            =   "frm_trans_bpjs.frx":05A6
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
      Height          =   7155
      Left            =   60
      TabIndex        =   11
      Top             =   690
      Width           =   14025
      _ExtentX        =   24739
      _ExtentY        =   12621
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "GENERAL"
      TabPicture(0)   =   "frm_trans_bpjs.frx":1638
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBGrid1(0)"
      Tab(0).Control(1)=   "fra_entry(0)"
      Tab(0).Control(2)=   "frmTombol(0)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "DEFAULT"
      TabPicture(1)   =   "frm_trans_bpjs.frx":1654
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TDBGrid1(1)"
      Tab(1).Control(1)=   "frmTombol(1)"
      Tab(1).Control(2)=   "fra_entry(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "ADMINISTRATION"
      TabPicture(2)   =   "frm_trans_bpjs.frx":1670
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label26(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lbl_department(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lbl_employee"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "TDBGrid1(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "fra_entry(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "TDBGrid_Emp"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "TDBCombo_department(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "TDBCombo_company(2)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "frmTombol(2)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txt_company_name(2)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txt_department_name(2)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Frame2"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "PROCCESS/REPORT"
      TabPicture(3)   =   "frm_trans_bpjs.frx":168C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lbl_department(1)"
      Tab(3).Control(1)=   "Label26(1)"
      Tab(3).Control(2)=   "Label14(4)"
      Tab(3).Control(3)=   "Label30(1)"
      Tab(3).Control(4)=   "lbl_jml_peserta(3)"
      Tab(3).Control(5)=   "cmdSearch(3)"
      Tab(3).Control(6)=   "TDBGrid1(3)"
      Tab(3).Control(7)=   "TDBCombo_department(3)"
      Tab(3).Control(8)=   "TDBCombo_company(3)"
      Tab(3).Control(9)=   "DTPicker_Periode(3)"
      Tab(3).Control(10)=   "fra_entry(3)"
      Tab(3).Control(11)=   "txt_department_name(3)"
      Tab(3).Control(12)=   "txt_company_name(3)"
      Tab(3).Control(13)=   "frmTombol(3)"
      Tab(3).ControlCount=   14
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   11040
         TabIndex        =   91
         Top             =   900
         Width           =   2745
         Begin VB.OptionButton optEmpty 
            Caption         =   "NOT EMPTY"
            Height          =   225
            Index           =   1
            Left            =   180
            TabIndex        =   93
            Top             =   210
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optEmpty 
            Caption         =   "EMPTY"
            Height          =   225
            Index           =   0
            Left            =   1650
            TabIndex        =   92
            Top             =   210
            Width           =   1035
         End
      End
      Begin VB.Frame fra_entry 
         Height          =   1545
         Index           =   1
         Left            =   -73620
         TabIndex        =   80
         Top             =   3300
         Width           =   11295
         Begin VB.ComboBox cbo_multiplier 
            Height          =   315
            ItemData        =   "frm_trans_bpjs.frx":16A8
            Left            =   4020
            List            =   "frm_trans_bpjs.frx":16B5
            TabIndex        =   87
            Text            =   "Basic Salary"
            Top             =   990
            Width           =   1695
         End
         Begin VB.TextBox txt_branch_bpjs 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5820
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   84
            Top             =   630
            Width           =   3495
         End
         Begin VB.TextBox txt_company_bpjs 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   1
            Left            =   5820
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   81
            Top             =   270
            Width           =   3495
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_company_bpjs 
            Height          =   375
            Index           =   1
            Left            =   4020
            OleObjectBlob   =   "frm_trans_bpjs.frx":16D8
            TabIndex        =   82
            Top             =   270
            Width           =   1695
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_branch_bpjs 
            Height          =   375
            Left            =   4020
            OleObjectBlob   =   "frm_trans_bpjs.frx":3AF2
            TabIndex        =   85
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "MULTIPLIER"
            Height          =   195
            Index           =   8
            Left            =   2580
            TabIndex        =   88
            Top             =   1050
            Width           =   960
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "BRANCH OFFICE*"
            Height          =   195
            Index           =   7
            Left            =   2580
            TabIndex        =   86
            Top             =   690
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "COMPANY*"
            Height          =   195
            Index           =   1
            Left            =   2580
            TabIndex        =   83
            Top             =   300
            Width           =   855
         End
      End
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1065
         Index           =   1
         Left            =   -73620
         TabIndex        =   74
         Top             =   4950
         Width           =   11295
         Begin prj_genting.vbButton cmdNew 
            Height          =   705
            Index           =   1
            Left            =   540
            TabIndex        =   75
            Top             =   240
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
            MICON           =   "frm_trans_bpjs.frx":5EF4
            PICN            =   "frm_trans_bpjs.frx":5F10
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdSave 
            Height          =   705
            Index           =   1
            Left            =   1560
            TabIndex        =   76
            Top             =   240
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
            MICON           =   "frm_trans_bpjs.frx":6FA2
            PICN            =   "frm_trans_bpjs.frx":6FBE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdEdit 
            Height          =   705
            Index           =   1
            Left            =   2580
            TabIndex        =   77
            Top             =   240
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
            MICON           =   "frm_trans_bpjs.frx":8050
            PICN            =   "frm_trans_bpjs.frx":806C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdDelete 
            Height          =   705
            Index           =   1
            Left            =   3600
            TabIndex        =   78
            Top             =   240
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
            MICON           =   "frm_trans_bpjs.frx":90FE
            PICN            =   "frm_trans_bpjs.frx":911A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdCancel 
            Height          =   705
            Index           =   1
            Left            =   4620
            TabIndex        =   79
            Top             =   240
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
            MICON           =   "frm_trans_bpjs.frx":A1AC
            PICN            =   "frm_trans_bpjs.frx":A1C8
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
      Begin VB.TextBox txt_department_name 
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
         Index           =   2
         Left            =   3420
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   810
         Width           =   2715
      End
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Index           =   2
         Left            =   3420
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   66
         Top             =   450
         Width           =   2715
      End
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1335
         Index           =   2
         Left            =   180
         TabIndex        =   60
         Top             =   5670
         Width           =   13605
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_genting.vbButton cmdNew 
            Height          =   705
            Index           =   2
            Left            =   540
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
            MICON           =   "frm_trans_bpjs.frx":B25A
            PICN            =   "frm_trans_bpjs.frx":B276
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdSave 
            Height          =   705
            Index           =   2
            Left            =   1560
            TabIndex        =   62
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
            MICON           =   "frm_trans_bpjs.frx":C308
            PICN            =   "frm_trans_bpjs.frx":C324
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdEdit 
            Height          =   705
            Index           =   2
            Left            =   2580
            TabIndex        =   63
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
            MICON           =   "frm_trans_bpjs.frx":D3B6
            PICN            =   "frm_trans_bpjs.frx":D3D2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdDelete 
            Height          =   705
            Index           =   2
            Left            =   3600
            TabIndex        =   64
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
            MICON           =   "frm_trans_bpjs.frx":E464
            PICN            =   "frm_trans_bpjs.frx":E480
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdCancel 
            Height          =   705
            Index           =   2
            Left            =   4620
            TabIndex        =   65
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
            MICON           =   "frm_trans_bpjs.frx":F512
            PICN            =   "frm_trans_bpjs.frx":F52E
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
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1335
         Index           =   3
         Left            =   -74820
         TabIndex        =   28
         Top             =   5670
         Width           =   13605
         Begin VB.Timer Timer3 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_genting.vbButton cmdPrint 
            Height          =   855
            Left            =   8760
            TabIndex        =   29
            Top             =   270
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1508
            BTYPE           =   14
            TX              =   "&Print"
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
            MICON           =   "frm_trans_bpjs.frx":105C0
            PICN            =   "frm_trans_bpjs.frx":105DC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdSave 
            Height          =   705
            Index           =   3
            Left            =   630
            TabIndex        =   30
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
            MICON           =   "frm_trans_bpjs.frx":1166E
            PICN            =   "frm_trans_bpjs.frx":1168A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdEdit 
            Height          =   705
            Index           =   3
            Left            =   1650
            TabIndex        =   31
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
            MICON           =   "frm_trans_bpjs.frx":1271C
            PICN            =   "frm_trans_bpjs.frx":12738
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdCancel 
            Height          =   705
            Index           =   3
            Left            =   2670
            TabIndex        =   32
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
            MICON           =   "frm_trans_bpjs.frx":137CA
            PICN            =   "frm_trans_bpjs.frx":137E6
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
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Index           =   3
         Left            =   -71580
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   27
         Top             =   450
         Width           =   2715
      End
      Begin VB.TextBox txt_department_name 
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
         Index           =   3
         Left            =   -71580
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   810
         Width           =   2715
      End
      Begin VB.Frame fra_entry 
         Height          =   2385
         Index           =   3
         Left            =   -74820
         TabIndex        =   21
         Top             =   3180
         Width           =   13605
         Begin VB.TextBox txt_description_proses 
            Height          =   735
            Left            =   5400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   810
            Width           =   4545
         End
         Begin VB.TextBox txt_bpjs 
            Height          =   315
            Left            =   10020
            TabIndex        =   22
            Top             =   810
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "DESCRIPTION"
            Height          =   165
            Left            =   2340
            TabIndex        =   24
            Top             =   1050
            Width           =   2895
         End
      End
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1065
         Index           =   0
         Left            =   -73620
         TabIndex        =   18
         Top             =   4770
         Width           =   11295
         Begin VB.Timer timer1 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   240
         End
         Begin prj_genting.vbButton cmdNew 
            Height          =   705
            Index           =   0
            Left            =   540
            TabIndex        =   5
            Top             =   240
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
            MICON           =   "frm_trans_bpjs.frx":14878
            PICN            =   "frm_trans_bpjs.frx":14894
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdSave 
            Height          =   705
            Index           =   0
            Left            =   1560
            TabIndex        =   6
            Top             =   240
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
            MICON           =   "frm_trans_bpjs.frx":15926
            PICN            =   "frm_trans_bpjs.frx":15942
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdEdit 
            Height          =   705
            Index           =   0
            Left            =   2580
            TabIndex        =   7
            Top             =   240
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
            MICON           =   "frm_trans_bpjs.frx":169D4
            PICN            =   "frm_trans_bpjs.frx":169F0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdDelete 
            Height          =   705
            Index           =   0
            Left            =   3600
            TabIndex        =   8
            Top             =   240
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
            MICON           =   "frm_trans_bpjs.frx":17A82
            PICN            =   "frm_trans_bpjs.frx":17A9E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_genting.vbButton cmdCancel 
            Height          =   705
            Index           =   0
            Left            =   4620
            TabIndex        =   9
            Top             =   240
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
            MICON           =   "frm_trans_bpjs.frx":18B30
            PICN            =   "frm_trans_bpjs.frx":18B4C
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
         Height          =   1515
         Index           =   0
         Left            =   -73620
         TabIndex        =   12
         Top             =   3240
         Width           =   11295
         Begin VB.TextBox txt_max_bpjs 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   7620
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txt_bpjs_kary 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   4020
            MaxLength       =   10
            TabIndex        =   1
            Top             =   660
            Width           =   1695
         End
         Begin VB.TextBox txt_bpjs_pers 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   4020
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txt_anak_maks 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   7620
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "3"
            Top             =   660
            Width           =   1695
         End
         Begin VB.TextBox txt_company_bpjs 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   0
            Left            =   5820
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   13
            Top             =   270
            Width           =   3495
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_company_bpjs 
            Height          =   375
            Index           =   0
            Left            =   4020
            OleObjectBlob   =   "frm_trans_bpjs.frx":19BDE
            TabIndex        =   0
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "MAKSIMAL UPAH*"
            Height          =   195
            Index           =   1
            Left            =   6030
            TabIndex        =   95
            Top             =   1050
            Width           =   1380
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PERUSAHAAN (%)*"
            Height          =   195
            Index           =   0
            Left            =   2580
            TabIndex        =   17
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "KARYAWAN (%)*"
            Height          =   195
            Index           =   0
            Left            =   2580
            TabIndex        =   16
            Top             =   690
            Width           =   1245
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "MAKSIMAL ANAK*"
            Height          =   195
            Index           =   0
            Left            =   6030
            TabIndex        =   15
            Top             =   690
            Width           =   1365
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "COMPANY*"
            Height          =   195
            Index           =   0
            Left            =   2580
            TabIndex        =   14
            Top             =   300
            Width           =   855
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3735
         Index           =   0
         Left            =   -73620
         TabIndex        =   20
         Top             =   1020
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6588
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "COMPANY CODE"
         Columns(0).DataField=   "company_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "COMPANY NAME"
         Columns(1).DataField=   "company_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "KARYAWAN"
         Columns(2).DataField=   "kary_value"
         Columns(2).NumberFormat=   "General Number"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "PERUSAHAAN"
         Columns(3).DataField=   "pers_value"
         Columns(3).NumberFormat=   "General Number"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "MAKSIMAL ANAK"
         Columns(4).DataField=   "maks_anak"
         Columns(4).NumberFormat=   "General Number"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "MAKSIMAL UPAH"
         Columns(5).DataField=   "max_bpjs_salary"
         Columns(5).NumberFormat=   "Standard"
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
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=7594"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7514"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2540"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2461"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2593"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2514"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2566"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2487"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=3625"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=3545"
         Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
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
         Caption         =   "LIST OF BPJS KESEHATAN"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=1"
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
      Begin MSComCtl2.DTPicker DTPicker_Periode 
         Height          =   315
         Index           =   3
         Left            =   -73080
         TabIndex        =   25
         Top             =   1170
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   96468995
         CurrentDate     =   41332
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Index           =   3
         Left            =   -73080
         OleObjectBlob   =   "frm_trans_bpjs.frx":1BFF8
         TabIndex        =   33
         Top             =   450
         Width           =   1455
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_department 
         Height          =   375
         Index           =   3
         Left            =   -73080
         OleObjectBlob   =   "frm_trans_bpjs.frx":1E40D
         TabIndex        =   34
         Top             =   810
         Width           =   1455
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4005
         Index           =   3
         Left            =   -74820
         TabIndex        =   35
         Top             =   1560
         Width           =   13605
         _ExtentX        =   23998
         _ExtentY        =   7064
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO. BPJS"
         Columns(0).DataField=   "no_bpjs"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "EMPLOYEE CODE"
         Columns(1).DataField=   "employee_code"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NIK"
         Columns(2).DataField=   "nik"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "EMPLOYEE NAME"
         Columns(3).DataField=   "employee_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DEPARTMENT"
         Columns(4).DataField=   "department_name"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DIVISION"
         Columns(5).DataField=   "division_name"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "TGL. LAHIR"
         Columns(6).DataField=   "date_birth"
         Columns(6).NumberFormat=   "dd-MM-yyyy"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "GAJI POKOK"
         Columns(7).DataField=   "basic_salary"
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "PERUSAHAAN"
         Columns(8).DataField=   "pers_value"
         Columns(8).NumberFormat=   "Standard"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "KARYAWAN"
         Columns(9).DataField=   "kary_value"
         Columns(9).NumberFormat=   "Standard"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "DESCRIPTION"
         Columns(10).DataField=   "description"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   11
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=11"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2381"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1746"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1667"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=4233"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=4154"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2275"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2196"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=2275"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2196"
         Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(32)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=2355"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2275"
         Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=513"
         Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(39)=   "Column(7).Width=2540"
         Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2461"
         Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(44)=   "Column(8).Width=2540"
         Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2461"
         Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=514"
         Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(49)=   "Column(9).Width=2408"
         Splits(0)._ColumnProps(50)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(9)._WidthInPix=2328"
         Splits(0)._ColumnProps(52)=   "Column(9)._ColStyle=514"
         Splits(0)._ColumnProps(53)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(54)=   "Column(10).Width=4657"
         Splits(0)._ColumnProps(55)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(56)=   "Column(10)._WidthInPix=4577"
         Splits(0)._ColumnProps(57)=   "Column(10)._ColStyle=516"
         Splits(0)._ColumnProps(58)=   "Column(10).Order=11"
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
         Caption         =   "LIST OF BPJS KESEHATAN"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
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
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=82,.parent=13,.alignment=1"
         _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=79,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=80,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=81,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
         _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
         _StyleDefs(78)  =   "Named:id=33:Normal"
         _StyleDefs(79)  =   ":id=33,.parent=0"
         _StyleDefs(80)  =   "Named:id=34:Heading"
         _StyleDefs(81)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(82)  =   ":id=34,.wraptext=-1"
         _StyleDefs(83)  =   "Named:id=35:Footing"
         _StyleDefs(84)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(85)  =   "Named:id=36:Selected"
         _StyleDefs(86)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(87)  =   "Named:id=37:Caption"
         _StyleDefs(88)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(89)  =   "Named:id=38:HighlightRow"
         _StyleDefs(90)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(91)  =   "Named:id=39:EvenRow"
         _StyleDefs(92)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(93)  =   "Named:id=40:OddRow"
         _StyleDefs(94)  =   ":id=40,.parent=33"
         _StyleDefs(95)  =   "Named:id=41:RecordSelector"
         _StyleDefs(96)  =   ":id=41,.parent=34"
         _StyleDefs(97)  =   "Named:id=42:FilterBar"
         _StyleDefs(98)  =   ":id=42,.parent=33"
      End
      Begin prj_genting.vbButton cmdSearch 
         Height          =   375
         Index           =   3
         Left            =   -71580
         TabIndex        =   36
         Top             =   1170
         Width           =   975
         _extentx        =   2249
         _extenty        =   820
         btype           =   14
         tx              =   "&View"
         enab            =   -1
         font            =   "frm_trans_bpjs.frx":2081D
         coltype         =   1
         focusr          =   -1
         bcol            =   15790320
         bcolo           =   15790320
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "frm_trans_bpjs.frx":20849
         picn            =   "frm_trans_bpjs.frx":20867
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Index           =   2
         Left            =   1920
         OleObjectBlob   =   "frm_trans_bpjs.frx":218FB
         TabIndex        =   68
         Top             =   450
         Width           =   1455
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_department 
         Height          =   375
         Index           =   2
         Left            =   1920
         OleObjectBlob   =   "frm_trans_bpjs.frx":23D10
         TabIndex        =   69
         Top             =   810
         Width           =   1455
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3735
         Index           =   1
         Left            =   -73620
         TabIndex        =   73
         Top             =   1110
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6588
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "COMPANY CODE"
         Columns(0).DataField=   "company_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "COMPANY NAME"
         Columns(1).DataField=   "company_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "BRANCH CODE"
         Columns(2).DataField=   "branch_code"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "BRANCH NAME"
         Columns(3).DataField=   "branch_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "MULTIPLIER"
         Columns(4).DataField=   "int_multiplier"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "MULTIPLIER"
         Columns(5).DataField=   "multiplier"
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
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=6826"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6747"
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
         Splits(0)._ColumnProps(23)=   "Column(4).Width=397"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=318"
         Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(27)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=5265"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=5186"
         Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
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
         Caption         =   "LIST OF BPJS DEFAULT MULTIPLIER"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Emp 
         Height          =   1875
         Left            =   180
         TabIndex        =   89
         Top             =   1410
         Width           =   13605
         _ExtentX        =   23998
         _ExtentY        =   3307
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
         Columns(2).Caption=   "DEPARTMENT CODE"
         Columns(2).DataField=   "department_code"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "DEPARTMENT"
         Columns(3).DataField=   "department_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "EMPLOYEE CODE"
         Columns(4).DataField=   "employee_code"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "NO. EMP ID."
         Columns(5).DataField=   "nik"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "EMPLOYEE NAME"
         Columns(6).DataField=   "employee_name"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "JOB TITLE"
         Columns(7).DataField=   "title_name"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
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
         Splits(0)._ColumnProps(17)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(21)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(23)=   "Column(3).Width=5662"
         Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=5583"
         Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(28)=   "Column(4).Width=2037"
         Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=1958"
         Splits(0)._ColumnProps(31)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(33)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(34)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(36)=   "Column(5).Width=3651"
         Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=3572"
         Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(40)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(41)=   "Column(6).Width=7699"
         Splits(0)._ColumnProps(42)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(6)._WidthInPix=7620"
         Splits(0)._ColumnProps(44)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(45)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(46)=   "Column(7).Width=5980"
         Splits(0)._ColumnProps(47)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(7)._WidthInPix=5900"
         Splits(0)._ColumnProps(49)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(50)=   "Column(7).Order=8"
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
         Caption         =   "LIST OF EMPLOYEE"
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
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=78,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=82,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=79,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=80,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=81,.parent=17"
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
      Begin VB.Frame fra_entry 
         Height          =   2385
         Index           =   2
         Left            =   180
         TabIndex        =   42
         Top             =   3270
         Width           =   13605
         Begin VB.CheckBox chkActive 
            Caption         =   "ACTIVE"
            Height          =   195
            Left            =   1980
            TabIndex        =   45
            Top             =   870
            Value           =   1  'Checked
            Width           =   1395
         End
         Begin VB.Frame Frame1 
            Caption         =   "ANAK"
            Height          =   1905
            Left            =   7200
            TabIndex        =   46
            Top             =   390
            Width           =   6285
            Begin VB.ListBox lst_anak 
               Appearance      =   0  'Flat
               Height          =   1590
               Left            =   3120
               TabIndex        =   50
               Top             =   150
               Width           =   3075
            End
            Begin VB.ComboBox cbo_anak 
               Height          =   315
               Left            =   120
               TabIndex        =   48
               Text            =   "Combo1"
               Top             =   780
               Width           =   2565
            End
            Begin prj_genting.vbButton cmdSet 
               Height          =   285
               Left            =   2730
               TabIndex        =   52
               Top             =   600
               Width           =   315
               _extentx        =   556
               _extenty        =   503
               btype           =   14
               tx              =   ">>"
               enab            =   -1
               font            =   "frm_trans_bpjs.frx":26120
               coltype         =   1
               focusr          =   -1
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_trans_bpjs.frx":2614C
               umcol           =   -1
               soft            =   0
               picpos          =   0
               ngrey           =   0
               fx              =   0
               hand            =   0
               check           =   0
               value           =   0
            End
            Begin prj_genting.vbButton cmdUnset 
               Height          =   285
               Left            =   2730
               TabIndex        =   53
               Top             =   930
               Width           =   315
               _extentx        =   556
               _extenty        =   503
               btype           =   14
               tx              =   "<<"
               enab            =   -1
               font            =   "frm_trans_bpjs.frx":2616A
               coltype         =   1
               focusr          =   -1
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_trans_bpjs.frx":26196
               umcol           =   -1
               soft            =   0
               picpos          =   0
               ngrey           =   0
               fx              =   0
               hand            =   0
               check           =   0
               value           =   0
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "ANAK :"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   54
               Top             =   540
               Width           =   525
            End
         End
         Begin VB.TextBox txt_no_bpjs 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   44
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txt_jml_anak 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   315
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   47
            Text            =   "0"
            Top             =   1140
            Width           =   1695
         End
         Begin VB.TextBox txt_nm_anak 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   405
            Left            =   1980
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   1470
            Width           =   4425
         End
         Begin VB.TextBox txt_description 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   1980
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   1890
            Width           =   4425
         End
         Begin MSComCtl2.DTPicker DTPicker_bpjs 
            Height          =   315
            Left            =   1980
            TabIndex        =   43
            Top             =   150
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   96468995
            CurrentDate     =   41332
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "STATUS"
            Height          =   195
            Index           =   2
            Left            =   1170
            TabIndex        =   90
            Top             =   870
            Width           =   645
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DATE"
            Height          =   195
            Index           =   1
            Left            =   1395
            TabIndex        =   59
            Top             =   210
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NO. BPJS*"
            Height          =   195
            Index           =   1
            Left            =   1050
            TabIndex        =   58
            Top             =   540
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JUMLAH ANAK"
            Height          =   195
            Index           =   2
            Left            =   735
            TabIndex        =   57
            Top             =   1200
            Width           =   1125
         End
         Begin VB.Label Label67 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA ANAK"
            Height          =   195
            Index           =   1
            Left            =   915
            TabIndex        =   56
            Top             =   1530
            Width           =   945
         End
         Begin VB.Label Label67 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Index           =   5
            Left            =   735
            TabIndex        =   55
            Top             =   1980
            Width           =   1095
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   2385
         Index           =   2
         Left            =   180
         TabIndex        =   70
         Top             =   3270
         Width           =   13605
         _ExtentX        =   23998
         _ExtentY        =   4207
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO. BPJS"
         Columns(0).DataField=   "no_bpjs"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "TGL. TERDAFTAR"
         Columns(1).DataField=   "reg_date"
         Columns(1).NumberFormat=   "dd-MM-yyyy"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "JML. ANAK"
         Columns(2).DataField=   "jml_anak"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NAMA ANAK"
         Columns(3).DataField=   "nm_anak"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   4
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "ACTIVE?"
         Columns(4).DataField=   "flag_active"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4286"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4207"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3175"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3096"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2143"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2064"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=4233"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=4154"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1588"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1508"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=7541"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=7461"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
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
         Caption         =   "LIST OF ADMINISTRATION"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=70,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=74,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=17"
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
      Begin VB.Label lbl_employee 
         Caption         =   "Total Employee"
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
         Left            =   240
         TabIndex        =   94
         Top             =   1170
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.Label lbl_department 
         AutoSize        =   -1  'True
         Caption         =   "DEPARTMENT"
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   72
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "COMPANY"
         Height          =   195
         Index           =   0
         Left            =   990
         TabIndex        =   71
         Top             =   510
         Width           =   795
      End
      Begin VB.Label lbl_jml_peserta 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   -66540
         TabIndex        =   41
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "JUMLAH PESERTA"
         Height          =   195
         Index           =   1
         Left            =   -68130
         TabIndex        =   40
         Top             =   1230
         Width           =   1440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "PERIODE"
         Height          =   195
         Index           =   4
         Left            =   -73920
         TabIndex        =   39
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "COMPANY"
         Height          =   195
         Index           =   1
         Left            =   -74010
         TabIndex        =   38
         Top             =   510
         Width           =   795
      End
      Begin VB.Label lbl_department 
         AutoSize        =   -1  'True
         Caption         =   "DEPARTMENT"
         Height          =   195
         Index           =   1
         Left            =   -74370
         TabIndex        =   37
         Top             =   870
         Width           =   1125
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "BPJS KESEHATAN"
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
      Left            =   360
      TabIndex        =   10
      Top             =   150
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_bpjs.frx":261B4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14760
   End
End
Attribute VB_Name = "frm_trans_bpjs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim SQL As String

Dim rsBPJS As New ADODB.Recordset
Dim rsBPJSClass As New ADODB.Recordset
Dim rsBPJSMultiplier As New ADODB.Recordset
Dim rsCompanyBPJS As New ADODB.Recordset
Dim rsBranchBPJS As New ADODB.Recordset
Dim rsAdmBPJS As New ADODB.Recordset
Dim rsProcessBPJS As New ADODB.Recordset

Dim rscompany As New ADODB.Recordset
Dim rsBranch As New ADODB.Recordset
Dim rsDepartment As New ADODB.Recordset
Dim rsDivision As New ADODB.Recordset

Dim rsemployee As New ADODB.Recordset
Dim rsCountEmployee As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Public public_int_mode As Integer
Dim vParam As String

Dim vSetData As Integer
Dim oClause As String
Dim x As Integer
Dim vSubject As String

Private Function check_validate_exist_new() As Boolean
    check_validate_exist_new = False
    
    If rs.State Then rs.Close
    If SSTab1.Tab = 0 Then
        SQL = "SELECT count(company_code) as rec_count FROM m_bpjs " & _
                  "WHERE company_code = '" & TDBCombo_company_bpjs(SSTab1.Tab).Columns("company_code").Value & "'"
    ElseIf SSTab1.Tab = 1 Then
        SQL = "SELECT count(branch_code) as rec_count FROM t_bpjs_multiplier " & _
                  "WHERE branch_code = '" & TDBCombo_branch_bpjs.Columns("branch_code").Value & "'"
    ElseIf SSTab1.Tab = 2 Then
        SQL = "SELECT count(no_bpjs) as rec_count FROM t_bpjs " & _
                  "WHERE no_bpjs = '" & txt_no_bpjs.Text & "'"
    End If
    rs.Open SQL, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        Exit Function
    End If
End Function

Private Sub check_invalid()
    MsgBox "Data sudah ada...", vbCritical, headerMSG
    
    If SSTab1.Tab = 0 Then
        TDBCombo_company_bpjs(SSTab1.Tab).Text = ""
        txt_company_bpjs(SSTab1.Tab).Text = ""
    ElseIf SSTab1.Tab = 1 Then
        TDBCombo_branch_bpjs.Text = ""
    ElseIf SSTab1.Tab = 2 Then
        txt_no_bpjs.Text = ""
    End If
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If SSTab1.Tab = 0 Then
        If Not TDBCombo_company_bpjs(SSTab1.Tab).Columns("company_code").Value = rsBPJS.Fields("company_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        If Not TDBCombo_branch_bpjs.Columns("branch_code").Value = rsBPJSMultiplier.Fields("branch_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        If Not txt_no_bpjs.Text = rsAdmBPJS.Fields("no_bpjs").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True
    
    If SSTab1.Tab = 0 Then
        'validasi BPJS Karyawan
        If Trim(txt_bpjs_kary) = "" Then
            MsgBox "Input BPJS Kesehatan Karyawan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_bpjs_kary.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi BPJS Perusahaan
        If Trim(txt_bpjs_pers) = "" Then
            MsgBox "Input BPJS Kesehatan Perusahaan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_bpjs_pers.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi Maksimal Anak
        If Trim(txt_anak_maks) = "" Then
            MsgBox "Input Maksimal Anak Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_anak_maks.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi Company
        If Trim(TDBCombo_company_bpjs(SSTab1.Tab)) = "" Then
            MsgBox "Input Perusahaan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_company_bpjs(SSTab1.Tab).SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        'validasi Company
        If Trim(TDBCombo_company_bpjs(SSTab1.Tab)) = "" Then
            MsgBox "Input Perusahaan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_company_bpjs(SSTab1.Tab).SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi Branch
        If Trim(TDBCombo_branch_bpjs) = "" Then
            MsgBox "Input Branch Office Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_branch_bpjs.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        If chkActive.Value = 1 Then
            'validasi No BPJS
            If Trim(txt_no_bpjs) = "" Then
                MsgBox "Input No BPJS Masih Kosong...", vbOKOnly + vbInformation, headerMSG
                txt_no_bpjs.SetFocus
                check_validate_new = False
                Exit Function
            End If
        End If
    End If
End Function

Private Sub chkActive_Click()
    If chkActive.Value = 0 Then
        Frame1.Enabled = False
    Else
        Frame1.Enabled = True
    End If
End Sub

Private Sub CmdCancel_Click(index As Integer)
    int_mode = 0
    Call load_mode
End Sub

Private Sub cmdDelete_Click(index As Integer)
Dim i As Integer
Dim vSalaryDate As String

On Error GoTo Err
    If Not (TDBGrid1(SSTab1.Tab).ApproxCount > 0 And TDBGrid1(SSTab1.Tab).Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
    CnG.BeginTrans
    If SSTab1.Tab = 0 Then
        i = MsgBox("Are you sure want to delete data '" _
            & TDBGrid1(SSTab1.Tab).Columns("company_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
                
        CnG.Execute "DELETE FROM m_bpjs WHERE company_code = '" & TDBGrid1(SSTab1.Tab).Columns("company_code").Value & "'"
    ElseIf SSTab1.Tab = 1 Then
        i = MsgBox("Are you sure want to delete data '" _
            & TDBGrid1(SSTab1.Tab).Columns("branch_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
                
        CnG.Execute "DELETE FROM t_bpjs_multiplier WHERE branch_code = '" & TDBGrid1(SSTab1.Tab).Columns("branch_code").Value & "'"
    ElseIf SSTab1.Tab = 2 Then
        If rs.State Then rs.Close
        SQL = "SELECT reg_date FROM t_bpjs " & _
                "WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "' " & _
                "ORDER BY reg_date DESC LIMIT 1"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            vSalaryDate = Format(rs.Fields(0).Value, "yyyy-MM-dd")
            
            If vSalaryDate = Format(TDBGrid1(SSTab1.Tab).Columns("reg_date").Value, "yyyy-MM-dd") Then
                i = MsgBox("Are you sure want to delete data '" _
                    & Format(TDBGrid1(SSTab1.Tab).Columns("reg_date").Value, "dd-MM-yyyy") & "' ?", vbYesNo + vbQuestion, headerMSG)
                If Not i = vbYes Then Exit Sub
                        
                CnG.Execute "DELETE FROM t_bpjs " & _
                    "WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "' " & _
                        "AND DATE(reg_date) = '" & Format(TDBGrid1(SSTab1.Tab).Columns("reg_date").Value, "yyyy-MM-dd") & "'"
            Else
                MsgBox "Data Cannot Be Delete, Because Is Not Last Date...", vbExclamation, headerMSG
                CnG.RollbackTrans
                Exit Sub
            End If
        End If
        rs.Close
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
    
    If Not (TDBGrid1(SSTab1.Tab).ApproxCount > 0 And TDBGrid1(SSTab1.Tab).Bookmark > 0) Then
        MsgBox "Tidak ada data yang dipilih...", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        With rsBPJS
            txt_bpjs_kary.Text = .Fields("kary_value").Value
            txt_bpjs_pers.Text = .Fields("pers_value").Value
            txt_anak_maks.Text = .Fields("maks_anak").Value
            txt_max_bpjs.Text = FormatNumber(.Fields("max_bpjs_salary").Value)
            
            Call set_data_company(.Fields("company_code").Value)
        End With
    ElseIf SSTab1.Tab = 1 Then
        With rsBPJSMultiplier
            Call load_data_company_bpjs
            Call load_data_branch_office_bpjs
            
            Call set_data_company(.Fields("company_code").Value)
            Call set_data_branch(.Fields("branch_code").Value)
            cbo_multiplier.ListIndex = .Fields("int_multiplier").Value
        End With
    ElseIf SSTab1.Tab = 2 Then
        cbo_anak.Clear
        lst_anak.Clear
        SQL = "DELETE FROM temp_bpjs_anak"
        CnG.Execute SQL
        
        SQL = "INSERT INTO temp_bpjs_anak " & _
              "SELECT * FROM t_bpjs_anak"
        CnG.Execute SQL
        
        With rsAdmBPJS
            txt_no_bpjs.Text = .Fields("no_bpjs").Value
            DTPicker_bpjs.Value = .Fields("reg_date").Value
            chkActive.Value = .Fields("flag_active").Value
            txt_jml_anak.Text = IIf(IsNull(.Fields("jml_anak").Value), 0, .Fields("jml_anak").Value)
            txt_nm_anak.Text = IIf(IsNull(.Fields("nm_anak").Value), "", .Fields("nm_anak").Value)
            txt_description.Text = IIf(IsNull(.Fields("description").Value), "", .Fields("description").Value)
        End With
        
        If rs.State Then rs.Close
        SQL = "SELECT nm_anak FROM t_bpjs_anak WHERE no_bpjs = '" & TDBGrid1(SSTab1.Tab).Columns("no_bpjs").Value & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            While Not rs.EOF
                lst_anak.AddItem rs.Fields(0).Value
                
                rs.MoveNext
            Wend
        End If
        rs.Close
        
        'Load Data Anak
        cbo_anak.Clear
            
        If rs.State Then rs.Close
        SQL = "SELECT name FROM m_employee_fams " & _
              "WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "' " & _
                "AND relationship = 2 " & _
                "AND name NOT IN (SELECT nm_anak FROM temp_bpjs_anak WHERE no_bpjs = '" & txt_no_bpjs.Text & "')"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            While Not rs.EOF
                cbo_anak.AddItem rs.Fields(0).Value
                
                rs.MoveNext
            Wend
            
            cbo_anak.ListIndex = 0
        End If
        rs.Close
    ElseIf SSTab1.Tab = 3 Then
        With rsProcessBPJS
            txt_bpjs.Text = TDBGrid1(SSTab1.Tab).Columns("no_bpjs").Value
            txt_description_proses.Text = IIf(IsNull(.Fields("description")), "", .Fields("description"))
        End With
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdEdit_Click(index As Integer)
    int_mode = 2
    Call load_mode
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdNew_Click(index As Integer)
    int_mode = 1
    Call load_mode
End Sub

Private Sub insert_new_data()
Dim i As Integer
Dim vInsurance As String

On Error GoTo Err
    CnG.BeginTrans
    
    If SSTab1.Tab = 0 Then
        SQL = "INSERT INTO m_bpjs(company_code,kary_value,pers_value,maks_anak,max_bpjs_salary) " _
                & "VALUES " _
                & "('" & TDBCombo_company_bpjs(SSTab1.Tab).Columns("company_code").Value & "','" & Trim(txt_bpjs_kary.Text) & "'," _
                & "'" & Trim(txt_bpjs_pers.Text) & "','" & Trim(txt_anak_maks.Text) & "','" & Val(DropAllComma(txt_max_bpjs.Text)) & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "INSERT INTO t_bpjs_multiplier(branch_code, int_multiplier, entry_date, entry_user) " _
                & "VALUES " _
                & "('" & TDBCombo_branch_bpjs.Columns("branch_code").Value & "','" & cbo_multiplier.ListIndex & "'," _
                & "Now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
'        If rs.State Then rs.Close
'        SQL = "SELECT b.insurance_name FROM t_insurance a JOIN m_insurance b ON a.insurance_code = b.insurance_code " & _
'              "WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "'"
'        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'        If rs.RecordCount > 0 Then
'            vInsurance = rs.Fields(0).Value
'
'            i = MsgBox("This Employee has followed insurance on '" & vInsurance & "'..." & _
'                        vbCrLf & "Do you want to continue?", vbYesNo + vbQuestion, headerMSG)
'            If Not i = vbYes Then Exit Sub
'        End If
'        rs.Close
        
        If chkActive.Value = 1 Then
            SQL = "INSERT INTO t_bpjs(no_bpjs,employee_code,reg_date,jml_anak,nm_anak,flag_active,description,entry_date,entry_user) " & _
                  "VALUES (" & _
                  "'" & txt_no_bpjs.Text & "','" & TDBGrid_Emp.Columns("employee_code").Value & "','" & Format(DTPicker_bpjs.Value, "yyyy-MM-dd") & "'," & _
                  "'" & txt_jml_anak.Text & "','" & txt_nm_anak.Text & "','" & chkActive.Value & "','" & txt_description.Text & "'," & _
                  "Now(),'" & LOGIN_NAME & "')"
        Else
            SQL = "INSERT INTO t_bpjs(no_bpjs,employee_code,reg_date,flag_active,description,entry_date,entry_user) " & _
                  "VALUES (" & _
                  "'" & txt_no_bpjs.Text & "','" & TDBGrid_Emp.Columns("employee_code").Value & "','" & Format(DTPicker_bpjs.Value, "yyyy-MM-dd") & "'," & _
                  "'" & chkActive.Value & "','" & txt_description.Text & "',Now(),'" & LOGIN_NAME & "')"
        End If
        CnG.Execute SQL
        
        SQL = "INSERT INTO t_bpjs_anak " & _
              "SELECT * FROM temp_bpjs_anak"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 3 Then
        SQL = "DELETE FROM t_bpjs_desc WHERE no_bpjs = '" & txt_bpjs.Text & "' AND month = '" & Format(DTPicker_Periode(SSTab1.Tab).Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
        
        SQL = "INSERT INTO t_bpjs_desc(no_bpjs,month,description,entry_date,entry_user) " & _
              "'" & txt_bpjs.Text & "','" & Format(DTPicker_Periode(SSTab1.Tab).Value, "yyyy-MM-dd") & "'," & _
              "'" & txt_description_proses.Text & "',Now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    End If

    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
On Error Resume Next

    CnG.BeginTrans
    
    If SSTab1.Tab = 0 Then
        SQL = "UPDATE m_bpjs SET company_code = '" & Trim(TDBCombo_company_bpjs(SSTab1.Tab).Columns("company_code").Value) & "'," & _
                "kary_value = '" & Trim(txt_bpjs_kary.Text) & "'," & _
                "pers_value = '" & Trim(txt_bpjs_pers.Text) & "'," & _
                "maks_anak = '" & Trim(txt_anak_maks.Text) & "'," & _
                "max_bpjs_salary = '" & Val(DropAllComma(txt_max_bpjs.Text)) & "' " & _
              "WHERE company_code = '" & TDBCombo_company_bpjs(SSTab1.Tab).Columns("company_code").Value & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "UPDATE t_bpjs_multiplier SET int_multiplier = '" & cbo_multiplier.ListIndex & "'," & _
                "edit_date = Now(), edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE branch_code = '" & TDBCombo_branch_bpjs.Columns("branch_code").Value & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "UPDATE t_bpjs SET no_bpjs = '" & txt_no_bpjs.Text & "'," & _
                "reg_date = '" & Format(DTPicker_bpjs.Value, "yyyy-MM-dd") & "'," & _
                "jml_anak = '" & txt_jml_anak.Text & "'," & _
                "nm_anak = '" & txt_nm_anak.Text & "'," & _
                "flag_active = '" & chkActive.Value & "'," & _
                "description = '" & txt_description.Text & "' " & _
              "WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "' " & _
                "AND DATE(reg_date) = '" & Format(DTPicker_bpjs.Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
        
        SQL = "INSERT INTO t_bpjs_anak " & _
              "SELECT * FROM temp_bpjs_anak"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 3 Then
        SQL = "DELETE FROM t_bpjs_desc WHERE no_bpjs = '" & txt_bpjs.Text & "' AND month = '" & Format(DTPicker_Periode(SSTab1.Tab).Value, "yyyy-MM") & "'"
        CnG.Execute SQL
        
        SQL = "INSERT INTO t_bpjs_desc(no_bpjs,month,description,entry_date,entry_user) " & _
              "VALUES('" & txt_bpjs.Text & "','" & Format(DTPicker_Periode(SSTab1.Tab).Value, "yyyy-MM") & "'," & _
              "'" & txt_description_proses.Text & "',Now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    End If
    
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdPrint_Click()
    With TDBGrid1(SSTab1.Tab).PrintInfo
        ' Set the page header.
        .PageHeaderFont.Italic = True
        .PageHeader = "List Of BPJS Kesehatan"

        ' Column headers will be on every page.
        .RepeatColumnHeaders = False

        ' Display page numbers (centered).
        .PageFooter = "\tPage: \p"

        ' Invoke Print Preview.
'        .SettingsPaperSize = 2
        .SettingsMarginLeft = 250
        .SettingsMarginRight = 150
        .SettingsMarginTop = 250
        .SettingsMarginBottom = 150
        .SettingsOrientation = 2
        .PreviewMaximize = True
        .PrintPreview
    End With
End Sub

Private Sub CmdSave_Click(index As Integer)
    If int_mode = 1 Then
        If Not check_validate_new Then Exit Sub
        If check_validate_exist_new Then
            Call check_invalid: Exit Sub
        End If
        Call insert_new_data
    ElseIf int_mode = 2 Then
'        If Not check_validate_new Then Exit Sub
'        If check_validate_exist_edit Then
'            Call check_invalid: Exit Sub
'        End If
        Call edit_old_data
    End If
    
    Call load_data
    int_mode = 0
    Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal F As Boolean, ByVal g As Boolean)
    If SSTab1.Tab <> 3 Then
        cmdNew(SSTab1.Tab).Enabled = a And blnUser_Add
        cmdSave(SSTab1.Tab).Enabled = b
        cmdEdit(SSTab1.Tab).Enabled = c And blnUser_Edit
        cmdDelete(SSTab1.Tab).Enabled = d And blnUser_Delete
        cmdCancel(SSTab1.Tab).Enabled = e
    Else
        cmdSave(SSTab1.Tab).Enabled = b
        cmdEdit(SSTab1.Tab).Enabled = c And blnUser_Edit
        cmdCancel(SSTab1.Tab).Enabled = e
    End If
End Sub

Private Sub clear_view_data()
Dim Ctr As Control
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_anak_maks" And _
               Not LCase(Ctr.name) = "txt_jml_anak" And _
               Not LCase(Ctr.name) = "txt_company_name" And _
               Not LCase(Ctr.name) = "txt_branch_name" And _
               Not LCase(Ctr.name) = "txt_department_name" And _
               Not LCase(Ctr.name) = "txt_division_name" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" And _
               Not LCase(Ctr.name) = "tdbcombo_branch" And _
               Not LCase(Ctr.name) = "tdbcombo_department" And _
               Not LCase(Ctr.name) = "tdbcombo_division" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is DTPicker Then
            If Not LCase(Ctr.name) = "dtpicker_periode" Then Ctr.Value = Now
        End If
    Next
End Sub

Private Sub set_enabled_control(ByVal i As Boolean)
Dim Ctr As Control
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
    If SSTab1.Tab = 1 Then
        cbo_multiplier.ListIndex = 0
    ElseIf SSTab1.Tab = 2 Then
        chkActive.Value = 1
        
        cbo_anak.Clear
        lst_anak.Clear
        
        SQL = "DELETE FROM temp_bpjs_anak"
        CnG.Execute SQL
    End If
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
                        
        If SSTab1.Tab = 0 Then
            fra_entry(SSTab1.Tab).Visible = True
            TDBCombo_company_bpjs(SSTab1.Tab).Enabled = True
            
            TDBGrid1(SSTab1.Tab).Enabled = False
            Call set_new_data
        ElseIf SSTab1.Tab = 1 Then
            fra_entry(SSTab1.Tab).Visible = True
            TDBCombo_company_bpjs(SSTab1.Tab).Enabled = True
            TDBCombo_branch_bpjs.Enabled = True
            
            TDBGrid1(SSTab1.Tab).Enabled = False
            Call set_new_data
        ElseIf SSTab1.Tab = 2 Then
            fra_entry(SSTab1.Tab).Visible = True
            DTPicker_bpjs.Enabled = True
            txt_no_bpjs.Enabled = True
            
            TDBGrid1(SSTab1.Tab).Enabled = False
            Call set_new_data
        End If
        
        If SSTab1.Tab = 0 Then
            If TDBCombo_company_bpjs(SSTab1.Tab).Enabled = True Then
                TDBCombo_company_bpjs(SSTab1.Tab).SetFocus
            End If
        ElseIf SSTab1.Tab = 2 Then
            If txt_no_bpjs.Enabled = True Then
                txt_no_bpjs.SetFocus
            End If
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        fra_entry(SSTab1.Tab).Visible = False
        TDBGrid1(SSTab1.Tab).Enabled = True
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        If SSTab1.Tab = 0 Then
            TDBCombo_company_bpjs(SSTab1.Tab).Enabled = False
        ElseIf SSTab1.Tab = 1 Then
            TDBCombo_company_bpjs(SSTab1.Tab).Enabled = False
            TDBCombo_branch_bpjs.Enabled = False
        ElseIf SSTab1.Tab = 2 Then
            DTPicker_bpjs.Enabled = False
            txt_no_bpjs.Enabled = False
        End If
        
        fra_entry(SSTab1.Tab).Visible = True
        TDBGrid1(SSTab1.Tab).Enabled = False
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
    SSTab1.TabVisible(3) = False
    
    Call load_data_company_bpjs
    oClause = ""
    
    SSTab1.Tab = 0
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    
    timer1.Enabled = True
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1(SSTab1.Tab).Columns
        Col.FilterText = ""
    Next Col
    
    If SSTab1.Tab = 0 Then
        rsBPJS.Filter = adFilterNone
    ElseIf SSTab1.Tab = 1 Then
        rsBPJSMultiplier.Filter = adFilterNone
    ElseIf SSTab1.Tab = 2 Then
        rsAdmBPJS.Filter = adFilterNone
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
    Set frm_trans_bpjs = Nothing
End Sub

Private Sub optEmpty_Click(index As Integer)
    Call load_data_employee
    Call load_count_data_employee
    
    Call load_data
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode

    If SSTab1.Tab = 0 Then
        Call load_data_company_bpjs
            
        timer1.Enabled = True
    ElseIf SSTab1.Tab = 1 Then
        Call load_data_company_bpjs
        Call load_data_branch_office_bpjs
            
        Call load_data
    ElseIf SSTab1.Tab = 2 Then
        Call load_data_company
        Call load_data_department
        
        Timer2.Enabled = True
    ElseIf SSTab1.Tab = 3 Then
        DTPicker_Periode(SSTab1.Tab).Value = Now
        
        Call load_data_company
        Call load_data_department
        
        Timer3.Enabled = True
    End If
End Sub

Private Sub TDBGrid_Emp_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If (TDBGrid_Emp.Row + 1) > 0 And (TDBGrid_Emp.Row + 1) <> LastRow Then
        'MsgBox "LETS..."
        Call load_data
    End If
End Sub

Private Sub TDBGrid1_FilterChange(index As Integer)
On Error GoTo Err

Dim i As Integer
    
    Set Cols = TDBGrid1(SSTab1.Tab).Columns
    i = TDBGrid1(SSTab1.Tab).Col
    TDBGrid1(SSTab1.Tab).HoldFields
    
    If SSTab1.Tab = 0 Then
        rsBPJS.Filter = getFilter()
    ElseIf SSTab1.Tab = 1 Then
        rsBPJSMultiplier.Filter = getFilter()
    ElseIf SSTab1.Tab = 2 Then
        rsAdmBPJS.Filter = getFilter()
    End If
    
    TDBGrid1(SSTab1.Tab).Col = i
    TDBGrid1(SSTab1.Tab).EditActive = True
    
    TDBGrid1(SSTab1.Tab).SelStart = Len(TDBGrid1(SSTab1.Tab).Columns(i).FilterText)
    If TDBGrid1(SSTab1.Tab).ApproxCount < 1 Then
        Call clear_filter
        TDBGrid1(SSTab1.Tab).Col = i
    End If
    
    Exit Sub
    
Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub TDBGrid_Emp_FilterChange()
On Error GoTo Err

Dim i As Integer

    Set Cols = TDBGrid_Emp.Columns
    i = TDBGrid_Emp.Col
    TDBGrid_Emp.HoldFields
    
    rsemployee.Filter = getFilter()
    TDBGrid_Emp.Col = i
    TDBGrid_Emp.EditActive = True
    
    TDBGrid_Emp.SelStart = Len(TDBGrid_Emp.Columns(i).FilterText)
    If TDBGrid_Emp.ApproxCount < 1 Then
        Call clear_filter_employee
        TDBGrid_Emp.Col = i
    End If
    
    Exit Sub
    
Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter_employee
End Sub


Private Sub load_data()
    If SSTab1.Tab = 0 Then
        If rsBPJS.State Then rsBPJS.Close
        SQL = "SELECT a.*,b.company_name " & _
              "FROM m_bpjs a LEFT JOIN m_company b on a.company_code = b.company_code " & oClause
        rsBPJS.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid1(SSTab1.Tab).DataSource = rsBPJS
    ElseIf SSTab1.Tab = 1 Then
        If rsBPJSMultiplier.State Then rsBPJSMultiplier.Close
        SQL = "SELECT a.*, c.company_code, c.company_name, b.branch_name, " & _
                "CASE WHEN a.int_multiplier = 0 THEN 'Basic Salary' WHEN a.int_multiplier = 1 THEN 'UMR' ELSE 'Premi BPJS' END multiplier " & _
              "FROM t_bpjs_multiplier a JOIN m_branch_office b on a.branch_code = b.branch_code " & _
              "JOIN m_company c ON b.company_code = c.company_code " & oClause
        rsBPJSMultiplier.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid1(SSTab1.Tab).DataSource = rsBPJSMultiplier
    ElseIf SSTab1.Tab = 2 Then
        vParam = IIf(TDBCombo_department(SSTab1.Tab).Text <> "", "b.department_code = '" & TDBCombo_department(SSTab1.Tab).Columns("department_code").Value & "' ", _
                    "b.company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' ")
                                
        If rsAdmBPJS.State Then rsAdmBPJS.Close
        SQL = "SELECT a.*,b.nik,b.employee_name,c.company_name,e.department_name," & _
                    "f.division_name " & _
              "FROM t_bpjs a JOIN m_employee b ON a.employee_code = b.employee_code " & _
              "LEFT JOIN m_company c ON b.company_code = c.company_code " & _
              "LEFT JOIN m_department e ON b.department_code = e.department_code " & _
              "LEFT JOIN m_division f ON b.division_code = f.division_code " & _
              "WHERE a.employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "' " & _
              "ORDER BY a.reg_date DESC"
        rsAdmBPJS.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid1(SSTab1.Tab).DataSource = rsAdmBPJS
    ElseIf SSTab1.Tab = 3 Then
        vParam = IIf(TDBCombo_department(SSTab1.Tab).Text <> "", "b.department_code = '" & TDBCombo_department(SSTab1.Tab).Columns("department_code").Value & "' ", _
                    "b.company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' ")
                                
        If rsProcessBPJS.State Then rsProcessBPJS.Close
        SQL = "SELECT a.no_bpjs,a.employee_code,b.date_birth,a.jml_anak,a.nm_anak," & _
                    "b.nik,b.employee_name,c.company_name,d.branch_name,e.department_name,f.division_name, " & _
                    "IFNULL((SELECT salary_value FROM h_salary WHERE employee_code = a.employee_code AND month = '" & Format(DTPicker_bpjs.Value, "yyyy-MM") & "' AND salary_code = 'SU-01'),0) basic_salary," & _
                    "IFNULL((SELECT salary_value FROM h_salary WHERE employee_code = a.employee_code AND month = '" & Format(DTPicker_bpjs.Value, "yyyy-MM") & "' AND salary_code = 'SU-038'),0) pers_value," & _
                    "IFNULL((SELECT salary_value FROM h_salary WHERE employee_code = a.employee_code AND month = '" & Format(DTPicker_bpjs.Value, "yyyy-MM") & "' AND salary_code = 'SU-282'),0) kary_value," & _
                    "g.description " & _
              "FROM t_bpjs a JOIN m_employee b ON a.employee_code = b.employee_code " & _
              "LEFT JOIN m_company c ON b.company_code = c.company_code " & _
              "LEFT JOIN m_department e ON b.department_code = e.department_code " & _
              "LEFT JOIN m_division f ON b.division_code = f.division_code " & _
              "LEFT JOIN t_bpjs_desc g ON a.no_bpjs = g.no_bpjs AND g.month = '" & Format(DTPicker_bpjs.Value, "yyyy-MM") & "' " & _
              "WHERE " & vParam & " " & _
                "AND LEFT(reg_date,7) = '" & Format(DTPicker_Periode(SSTab1.Tab).Value, "yyyy-MM") & "' " & oClause
        rsProcessBPJS.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid1(SSTab1.Tab).DataSource = rsProcessBPJS
    End If
End Sub

Private Sub load_data_count()
    vParam = IIf(TDBCombo_department(SSTab1.Tab).Text <> "", "b.department_code = '" & TDBCombo_department(SSTab1.Tab).Columns("department_code").Value & "' ", _
                    "b.company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' ")
                                
    If rs.State Then rs.Close
    SQL = "SELECT COUNT(a.employee_code) " & _
          "FROM t_bpjs a JOIN m_employee b ON a.employee_code = b.employee_code " & _
          "LEFT JOIN m_company c ON b.company_code = c.company_code " & _
          "LEFT JOIN m_department e ON b.department_code = e.department_code " & _
          "LEFT JOIN m_division f ON b.division_code = f.division_code " & _
          "WHERE " & vParam & " " & _
            "AND LEFT(reg_date,7) = '" & Format(DTPicker_Periode(SSTab1.Tab).Value, "yyyy-MM") & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        lbl_jml_peserta(SSTab1.Tab).Caption = IIf(IsNull(rs.Fields(0).Value), 0, rs.Fields(0).Value)
    Else
        lbl_jml_peserta(SSTab1.Tab).Caption = 0
    End If
    rs.Close
End Sub

Private Sub Timer1_Timer()
    Call load_data
    timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    If SSTab1.Tab <> 0 Then
        Call set_company_mode_rs(rscompany, TDBCombo_company(SSTab1.Tab), txt_company_name(SSTab1.Tab))
    End If
    
    If SSTab1.Tab = 2 Then
        Call load_data_employee
    End If
    
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    Call set_company_mode_rs(rscompany, TDBCombo_company(SSTab1.Tab), txt_company_name(SSTab1.Tab))
    
    Timer3.Enabled = False
End Sub

Private Sub TDBCombo_company_Change(index As Integer)
    If TDBCombo_company(SSTab1.Tab).Text = "" Then txt_company_name(SSTab1.Tab).Text = ""
End Sub

Private Sub TDBCombo_department_Change(index As Integer)
    If TDBCombo_department(SSTab1.Tab).Text = "" Then txt_department_name(SSTab1.Tab).Text = ""
End Sub

Private Sub TDBCombo_company_bpjs_ItemChange(index As Integer)
    If TDBCombo_company_bpjs(SSTab1.Tab).ApproxCount > 0 Then
        TDBCombo_company_bpjs(SSTab1.Tab).Text = TDBCombo_company_bpjs(SSTab1.Tab).Columns("company_code").Value
        txt_company_bpjs(SSTab1.Tab).Text = TDBCombo_company_bpjs(SSTab1.Tab).Columns("company_name").Value
        
        If SSTab1.Tab = 1 Then
            Call load_data_branch_office_bpjs
        End If
    End If
End Sub

Private Sub TDBCombo_branch_bpjs_ItemChange()
    If TDBCombo_branch_bpjs.ApproxCount > 0 Then
        TDBCombo_branch_bpjs.Text = TDBCombo_branch_bpjs.Columns("branch_code").Value
        txt_branch_bpjs.Text = TDBCombo_branch_bpjs.Columns("branch_name").Value
    End If
End Sub

Private Sub TDBCombo_company_ItemChange(index As Integer)
    If TDBCombo_company(SSTab1.Tab).ApproxCount > 0 Then
        TDBCombo_company(SSTab1.Tab).Text = TDBCombo_company(SSTab1.Tab).Columns("company_code").Value
        txt_company_name(SSTab1.Tab).Text = TDBCombo_company(SSTab1.Tab).Columns("company_name").Value

        Call load_data_department
        
        Call load_data_employee
        Call load_count_data_employee
        
        Call load_data
    End If
End Sub

Private Sub TDBCombo_department_ItemChange(index As Integer)
    If TDBCombo_department(SSTab1.Tab).ApproxCount > 0 Then
        TDBCombo_department(SSTab1.Tab).Text = TDBCombo_department(SSTab1.Tab).Columns("department_code").Value
        txt_department_name(SSTab1.Tab).Text = TDBCombo_department(SSTab1.Tab).Columns("department_name").Value
        
        Call load_data_employee
        Call load_count_data_employee
        
        Call load_data
    End If
End Sub

Private Sub load_data_company_bpjs()
    If rsCompanyBPJS.State Then rsCompanyBPJS.Close
    SQL = "select * from m_company order by company_code"
    rsCompanyBPJS.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Then
        TDBCombo_company_bpjs(SSTab1.Tab).RowSource = rsCompanyBPJS
    End If
End Sub

Public Sub load_data_branch_office_bpjs()
    TDBCombo_branch_bpjs.Text = "": txt_branch_bpjs.Text = ""
    
    If rsBranchBPJS.State Then rsBranchBPJS.Close
    SQL = "select * from m_branch_office " & _
          "where company_code = '" & TDBCombo_company_bpjs(SSTab1.Tab).Columns("company_code").Value & "' " & _
          "order by branch_code"
    rsBranchBPJS.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_branch_bpjs.RowSource = rsBranchBPJS
End Sub

Public Sub load_data_company()
    TDBCombo_company(SSTab1.Tab).Text = "": txt_company_name(SSTab1.Tab) = ""
    
    If rscompany.State Then rscompany.Close
    SQL = "select * from m_company order by company_code"
    rscompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company(SSTab1.Tab).RowSource = rscompany
End Sub

Public Sub load_data_department()
    TDBCombo_department(SSTab1.Tab).Text = "": txt_department_name(SSTab1.Tab) = ""

    If rsDepartment.State Then rsDepartment.Close
    SQL = "select * from m_department " & _
          "where company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' " & _
          "order by department_code"
    rsDepartment.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_department(SSTab1.Tab).RowSource = rsDepartment
End Sub

Private Sub TDBGrid1_HeadClick(index As Integer, ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid1(SSTab1.Tab).Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid1(SSTab1.Tab).Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid1(SSTab1.Tab).Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid1(SSTab1.Tab).Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub set_data_company(ByVal str_code As String)
On Error GoTo Err
    
    If rsCompanyBPJS.RecordCount > 0 Then
        rsCompanyBPJS.MoveFirst
        rsCompanyBPJS.Find ("company_code='" & str_code & "'")   ', 0, adSearchForward, 1)
        If Not (rsCompanyBPJS.EOF = True Or rsCompanyBPJS.BOF = True) Then
'            TDBCombo_company.Bookmark = rsCompanyBPJS.AbsolutePosition
            Call TDBCombo_company_bpjs_ItemChange(SSTab1.Tab)
        Else
            TDBCombo_company_bpjs(SSTab1.Tab).Text = ""
        End If
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub set_data_branch(ByVal str_code As String)
On Error GoTo Err
    
    If rsBranchBPJS.RecordCount > 0 Then
        rsBranchBPJS.MoveFirst
        rsBranchBPJS.Find ("branch_code='" & str_code & "'")   ', 0, adSearchForward, 1)
        If Not (rsBranchBPJS.EOF = True Or rsBranchBPJS.BOF = True) Then
'            TDBCombo_company.Bookmark = rsCompanyBPJS.AbsolutePosition
            Call TDBCombo_branch_bpjs_ItemChange
        Else
            TDBCombo_branch_bpjs.Text = ""
        End If
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdSet_Click()
Dim vMaxAnak As String
Dim i As Integer
   
On Error Resume Next
    If rs.State Then rs.Close
    SQL = "SELECT maks_anak FROM m_bpjs WHERE company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vMaxAnak = IIf(IsNull(rs.Fields(0).Value), 0, rs.Fields(0).Value)
    Else
        vMaxAnak = 0
    End If
    rs.Close
    
    If rs.State Then rs.Close
    SQL = "SELECT IFNULL(COUNT(nm_anak),0) FROM temp_bpjs_anak WHERE no_bpjs = '" & txt_no_bpjs.Text & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        If rs.Fields(0).Value >= vMaxAnak Then
            MsgBox "Batas Anak Sudah Maksimal...", vbExclamation, headerMSG
            Exit Sub
        Else
            SQL = "INSERT INTO temp_bpjs_anak(no_bpjs,nm_anak) " & _
                  "VALUES ('" & txt_no_bpjs.Text & "','" & cbo_anak.Text & "')"
            CnG.Execute SQL
            
            lst_anak.AddItem cbo_anak.Text
        End If
    End If
    rs.Close
    
    'Insert nama anak
    If rs.State Then rs.Close
    SQL = "SELECT nm_anak FROM temp_bpjs_anak WHERE no_bpjs = '" & txt_no_bpjs & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    i = 0
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            i = i + 1
            If i = 1 Then
                txt_nm_anak.Text = rs.Fields(0).Value
            Else
                txt_nm_anak.Text = txt_nm_anak.Text & "," & rs.Fields(0).Value
            End If
            
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    'Load Count Anak
    If rs.State Then rs.Close
    SQL = "SELECT IFNULL(COUNT(nm_anak),0) FROM temp_bpjs_anak WHERE no_bpjs = '" & txt_no_bpjs.Text & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        txt_jml_anak.Text = rs.Fields(0).Value
    End If
    rs.Close
    
    'Load Data Anak
    cbo_anak.Clear
        
    If rs.State Then rs.Close
    SQL = "SELECT name FROM m_employee_fams " & _
          "WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "' " & _
            "AND relationship = 2 " & _
            "AND name NOT IN (SELECT nm_anak FROM temp_bpjs_anak WHERE no_bpjs = '" & txt_no_bpjs.Text & "')"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            cbo_anak.AddItem rs.Fields(0).Value
            
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    cbo_anak.ListIndex = 0
End Sub

Private Sub cmdUnset_Click()
Dim vMaxAnak As String
Dim i As Integer
   
On Error Resume Next
    If rs.State Then rs.Close
    SQL = "SELECT maks_anak FROM m_bpjs WHERE company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vMaxAnak = IIf(IsNull(rs.Fields(0).Value), 0, rs.Fields(0).Value)
    Else
        vMaxAnak = 0
    End If
    rs.Close
    
    SQL = "DELETE FROM temp_bpjs_anak WHERE no_bpjs = '" & txt_no_bpjs.Text & "' " & _
            "AND nm_anak = '" & lst_anak.Text & "'"
    CnG.Execute SQL
    
    'Insert nama anak
    If rs.State Then rs.Close
    SQL = "SELECT nm_anak FROM temp_bpjs_anak WHERE no_bpjs = '" & txt_no_bpjs & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    i = 0
    txt_nm_anak.Text = ""
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            i = i + 1
            If i = 1 Then
                txt_nm_anak.Text = rs.Fields(0).Value
            Else
                txt_nm_anak.Text = txt_nm_anak.Text & "," & rs.Fields(0).Value
            End If
            
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    'Load Count Anak
    If rs.State Then rs.Close
    SQL = "SELECT IFNULL(COUNT(nm_anak),0) FROM temp_bpjs_anak WHERE no_bpjs = '" & txt_no_bpjs.Text & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        txt_jml_anak.Text = rs.Fields(0).Value
    End If
    rs.Close
    
    'Load list data anak
    lst_anak.Clear
    
    If rs.State Then rs.Close
    SQL = "SELECT nm_anak FROM temp_bpjs_anak WHERE no_bpjs = '" & txt_no_bpjs & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            lst_anak.AddItem rs.Fields(0).Value
            
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    'Load Data Anak
    cbo_anak.Clear
        
    If rs.State Then rs.Close
    SQL = "SELECT name FROM m_employee_fams " & _
          "WHERE employee_code = '" & TDBGrid_Emp.Columns("employe_code").Value & "' " & _
            "AND relationship = 2 " & _
            "AND name NOT IN (SELECT nm_anak FROM temp_bpjs_anak WHERE no_bpjs = '" & txt_no_bpjs.Text & "')"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            cbo_anak.AddItem rs.Fields(0).Value
            
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    cbo_anak.ListIndex = 0
End Sub

Public Sub load_data_employee()
                        
    If rsemployee.State Then rsemployee.Close
    If LOGIN_LEVEL = 100 Then
        SQL = "SELECT DISTINCT a.*,c.title_name,e.department_name " & _
              "FROM m_employee a " & _
                "JOIN m_title c ON a.title_code = c.title_code " & _
                "LEFT JOIN t_bpjs d ON a.employee_code = d.employee_code " & _
                "JOIN m_department e ON a.department_code = e.department_code " & _
              "WHERE CASE WHEN '" & TDBCombo_department(SSTab1.Tab).Text & "' = '' THEN " & _
                        "a.company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' " & _
                     "ELSE " & _
                        "a.company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' " & _
                        "AND a.department_code = '" & TDBCombo_department(SSTab1.Tab).Columns("department_code").Value & "' " & _
                     "END " & _
                "AND a.flag_active <> 0 " & _
                "AND " & IIf(optEmpty(0).Value, "IFNULL(d.no_bpjs,0) = 0 ", "IFNULL(d.no_bpjs,0) <> 0 ") & " " & oClause
    Else
        SQL = "SELECT DISTINCT a.*,c.title_name,e.department_name " & _
              "FROM m_employee a " & _
                "JOIN m_title c ON a.title_code = c.title_code " & _
                "LEFT JOIN t_bpjs d ON a.employee_code = d.employee_code " & _
                "JOIN m_department e ON a.department_code = e.department_code " & _
              "WHERE CASE WHEN '" & TDBCombo_department(SSTab1.Tab).Text & "' = '' THEN " & _
                        "a.company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' " & _
                     "ELSE " & _
                        "a.company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' " & _
                        "AND a.department_code = '" & TDBCombo_department(SSTab1.Tab).Columns("department_code").Value & "' " & _
                     "END " & _
                "AND (a.level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND a.flag_active <> 0 " & _
                "AND " & IIf(optEmpty(0).Value, "IFNULL(d.no_bpjs,0) = 0 ", "IFNULL(d.no_bpjs,0) <> 0 ") & " " & oClause
    End If

    rsemployee.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid_Emp.DataSource = rsemployee
End Sub

Public Sub load_count_data_employee()
                   
    If rsCountEmployee.State Then rsCountEmployee.Close
    If LOGIN_LEVEL = 100 Then
        SQL = "SELECT COUNT(DISTINCT a.employee_code) " & _
              "FROM m_employee a JOIN m_department c ON a.department_code = c.department_code " & _
                "JOIN m_title d ON a.title_code = d.title_code " & _
                "LEFT JOIN t_bpjs e ON a.employee_code = e.employee_code " & _
              "WHERE CASE WHEN '" & TDBCombo_department(SSTab1.Tab).Text & "' = '' THEN " & _
                        "a.company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' " & _
                     "ELSE " & _
                        "a.company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' " & _
                        "AND a.department_code = '" & TDBCombo_department(SSTab1.Tab).Columns("department_code").Value & "' " & _
                     "END " & _
                "AND a.flag_active <> 0 " & _
                "AND " & IIf(optEmpty(0).Value, "IFNULL(e.no_bpjs,0) = 0 ", "IFNULL(e.no_bpjs,0) <> 0 ") & " "
    Else
        SQL = "SELECT COUNT(DISTINCT a.employee_code) " & _
              "FROM m_employee a JOIN m_department c ON a.department_code = c.department_code " & _
                "JOIN m_title d ON a.title_code = d.title_code " & _
                "LEFT JOIN t_bpjs e ON a.employee_code = e.employee_code " & _
              "WHERE CASE WHEN '" & TDBCombo_department(SSTab1.Tab).Text & "' = '' THEN " & _
                        "a.company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' " & _
                     "ELSE " & _
                        "a.company_code = '" & TDBCombo_company(SSTab1.Tab).Columns("company_code").Value & "' " & _
                        "AND a.department_code = '" & TDBCombo_department(SSTab1.Tab).Columns("department_code").Value & "' " & _
                     "END " & _
                "AND (a.level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND a.flag_active <> 0 " & _
                "AND " & IIf(optEmpty(0).Value, "IFNULL(e.no_bpjs,0) = 0 ", "IFNULL(e.no_bpjs,0) <> 0 ") & " "
    End If

    rsCountEmployee.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    lbl_employee.Visible = True
    lbl_employee.Caption = "Total Employee : " & rsCountEmployee.Fields(0).Value
End Sub

Private Sub TDBGrid_Emp_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Emp.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Emp.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Emp.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Emp.Columns(ColIndex).DataField
    Call load_data_employee

End Sub

Private Sub clear_filter_employee()
    For Each Col In TDBGrid_Emp.Columns
        Col.FilterText = ""
    Next Col
    rsemployee.Filter = adFilterNone
End Sub

Private Sub txt_max_bpjs_Validate(Cancel As Boolean)
    If Not Trim(txt_max_bpjs.Text) = "" Then
        txt_max_bpjs.Text = FormatNumber(DropAllComma(txt_max_bpjs.Text))
    End If
End Sub
