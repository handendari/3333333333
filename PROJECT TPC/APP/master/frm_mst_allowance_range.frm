VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_allowance_range 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER ALLOWANCE RANGE"
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
   Icon            =   "frm_mst_allowance_range.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "frm_mst_allowance_range.frx":058A
   ScaleHeight     =   8955
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   90
      TabIndex        =   2
      Top             =   870
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "GENERAL"
      TabPicture(0)   =   "frm_mst_allowance_range.frx":0B14
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label27"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TDBGrid_Gen"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_entry_general"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TDBCombo_general"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txt_general"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_entry_dtl_general"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "TRAVELLING ALLOWANCE"
      TabPicture(1)   =   "frm_mst_allowance_range.frx":0B30
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_entry_Travel"
      Tab(1).Control(1)=   "frmTombol"
      Tab(1).Control(2)=   "TDBGrid_Travel"
      Tab(1).ControlCount=   3
      Begin VB.Frame fra_entry_Travel 
         Height          =   2655
         Left            =   -74730
         TabIndex        =   94
         Top             =   2730
         Width           =   11295
         Begin VB.TextBox txt_title_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5310
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   99
            Top             =   840
            Width           =   3855
         End
         Begin VB.TextBox txt_allow_travel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3840
            MaxLength       =   50
            TabIndex        =   102
            Top             =   1200
            Width           =   1395
         End
         Begin VB.TextBox txt_description_travel 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3840
            MaxLength       =   50
            TabIndex        =   103
            Top             =   1560
            Width           =   3495
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_title 
            Height          =   375
            Left            =   3840
            OleObjectBlob   =   "frm_mst_allowance_range.frx":0B4C
            TabIndex        =   100
            Top             =   840
            Width           =   1425
         End
         Begin VB.Label Label3 
            Caption         =   "USD"
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
            Left            =   5310
            TabIndex        =   101
            Top             =   1230
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "ALLOWANCE*"
            Height          =   195
            Left            =   2400
            TabIndex        =   97
            Top             =   1200
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "JOB TITLE*"
            Height          =   195
            Left            =   2400
            TabIndex        =   96
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   2400
            TabIndex        =   95
            Top             =   1560
            Width           =   1020
         End
      End
      Begin VB.Frame fra_entry_dtl_general 
         Height          =   2775
         Left            =   240
         TabIndex        =   79
         Top             =   2610
         Width           =   11295
         Begin VB.CheckBox chkUp 
            Height          =   255
            Left            =   5520
            TabIndex        =   82
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtCompensation_Gen 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5520
            MaxLength       =   50
            TabIndex        =   83
            Top             =   1560
            Width           =   1005
         End
         Begin VB.TextBox txtTo_Gen 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5520
            TabIndex        =   81
            Top             =   840
            Width           =   1005
         End
         Begin VB.TextBox txtDescription_Gen 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5520
            MaxLength       =   50
            TabIndex        =   84
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txtFrom_Gen 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5520
            TabIndex        =   80
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label16 
            Caption         =   "Years"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   6600
            TabIndex        =   93
            Top             =   870
            Width           =   495
         End
         Begin VB.Label Label15 
            Caption         =   "Years"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   6600
            TabIndex        =   92
            Top             =   540
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "x BASIC SALARY"
            Height          =   225
            Left            =   6600
            TabIndex        =   90
            Top             =   1620
            Width           =   1815
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "COMPENSATION*"
            Height          =   195
            Left            =   4200
            TabIndex        =   89
            Top             =   1560
            Width           =   1290
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "UP"
            Height          =   195
            Left            =   4200
            TabIndex        =   88
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "TO VALUE*"
            Height          =   195
            Left            =   4200
            TabIndex        =   87
            Top             =   840
            Width           =   810
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   4200
            TabIndex        =   86
            Top             =   1920
            Width           =   1020
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "FROM VALUE*"
            Height          =   195
            Left            =   4200
            TabIndex        =   85
            Top             =   480
            Width           =   1035
         End
      End
      Begin VB.TextBox txt_general 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   68
         Top             =   600
         Width           =   3855
      End
      Begin VB.Frame fra_entry_Department 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   48
         Top             =   2280
         Width           =   11205
         Begin VB.TextBox txt_department_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   50
            TabIndex        =   51
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txt_department_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   50
            TabIndex        =   50
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txt_department_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   10
            TabIndex        =   49
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   3480
            TabIndex        =   54
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   3480
            TabIndex        =   53
            Top             =   600
            Width           =   510
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   3480
            TabIndex        =   52
            Top             =   960
            Width           =   525
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   39
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
            TabIndex        =   40
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
            MICON           =   "frm_mst_allowance_range.frx":2AA8
            PICN            =   "frm_mst_allowance_range.frx":2AC4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_Department 
            Height          =   705
            Left            =   1560
            TabIndex        =   41
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
            MICON           =   "frm_mst_allowance_range.frx":3B56
            PICN            =   "frm_mst_allowance_range.frx":3B72
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_Department 
            Height          =   705
            Left            =   2580
            TabIndex        =   42
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
            MICON           =   "frm_mst_allowance_range.frx":4C04
            PICN            =   "frm_mst_allowance_range.frx":4C20
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_Department 
            Height          =   705
            Left            =   3600
            TabIndex        =   43
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
            MICON           =   "frm_mst_allowance_range.frx":5CB2
            PICN            =   "frm_mst_allowance_range.frx":5CCE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_Department 
            Height          =   705
            Left            =   4620
            TabIndex        =   44
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
            MICON           =   "frm_mst_allowance_range.frx":6D60
            PICN            =   "frm_mst_allowance_range.frx":6D7C
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
         TabIndex        =   29
         Top             =   1770
         Width           =   11175
         Begin VB.TextBox txt_department 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5400
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   33
            Top             =   660
            Width           =   3855
         End
         Begin VB.TextBox txt_division_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   32
            Top             =   1770
            Width           =   3495
         End
         Begin VB.TextBox txt_division_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   31
            Top             =   1410
            Width           =   3495
         End
         Begin VB.TextBox txt_division_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   30
            Top             =   1050
            Width           =   1695
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_department 
            Height          =   375
            Left            =   3600
            OleObjectBlob   =   "frm_mst_allowance_range.frx":7E0E
            TabIndex        =   34
            Top             =   660
            Width           =   1725
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "DEPARTMENT*"
            Height          =   195
            Left            =   2280
            TabIndex        =   38
            Top             =   690
            Width           =   1185
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   2280
            TabIndex        =   37
            Top             =   1770
            Width           =   1095
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   2280
            TabIndex        =   36
            Top             =   1050
            Width           =   510
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   2280
            TabIndex        =   35
            Top             =   1410
            Width           =   525
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   23
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
            TabIndex        =   24
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
            MICON           =   "frm_mst_allowance_range.frx":9D77
            PICN            =   "frm_mst_allowance_range.frx":9D93
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_Division 
            Height          =   705
            Left            =   1560
            TabIndex        =   25
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
            MICON           =   "frm_mst_allowance_range.frx":AE25
            PICN            =   "frm_mst_allowance_range.frx":AE41
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_Division 
            Height          =   705
            Left            =   2580
            TabIndex        =   26
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
            MICON           =   "frm_mst_allowance_range.frx":BED3
            PICN            =   "frm_mst_allowance_range.frx":BEEF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_Division 
            Height          =   705
            Left            =   3600
            TabIndex        =   27
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
            MICON           =   "frm_mst_allowance_range.frx":CF81
            PICN            =   "frm_mst_allowance_range.frx":CF9D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_Division 
            Height          =   705
            Left            =   4620
            TabIndex        =   28
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
            MICON           =   "frm_mst_allowance_range.frx":E02F
            PICN            =   "frm_mst_allowance_range.frx":E04B
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
         TabIndex        =   16
         Top             =   1770
         Width           =   11175
         Begin VB.TextBox txt_grade_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   10
            TabIndex        =   19
            Top             =   810
            Width           =   1695
         End
         Begin VB.TextBox txt_grade_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   50
            TabIndex        =   18
            Top             =   1170
            Width           =   3495
         End
         Begin VB.TextBox txt_grade_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   50
            TabIndex        =   17
            Top             =   1530
            Width           =   3495
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   2910
            TabIndex        =   22
            Top             =   1170
            Width           =   525
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   2910
            TabIndex        =   21
            Top             =   810
            Width           =   510
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   2910
            TabIndex        =   20
            Top             =   1530
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74790
         TabIndex        =   10
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
            TabIndex        =   11
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
            MICON           =   "frm_mst_allowance_range.frx":F0DD
            PICN            =   "frm_mst_allowance_range.frx":F0F9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_Grade 
            Height          =   705
            Left            =   1560
            TabIndex        =   12
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
            MICON           =   "frm_mst_allowance_range.frx":1018B
            PICN            =   "frm_mst_allowance_range.frx":101A7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_Grade 
            Height          =   705
            Left            =   2580
            TabIndex        =   13
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
            MICON           =   "frm_mst_allowance_range.frx":11239
            PICN            =   "frm_mst_allowance_range.frx":11255
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_Grade 
            Height          =   705
            Left            =   3600
            TabIndex        =   14
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
            MICON           =   "frm_mst_allowance_range.frx":122E7
            PICN            =   "frm_mst_allowance_range.frx":12303
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_Grade 
            Height          =   705
            Left            =   4620
            TabIndex        =   15
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
            MICON           =   "frm_mst_allowance_range.frx":13395
            PICN            =   "frm_mst_allowance_range.frx":133B1
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
         TabIndex        =   3
         Top             =   1650
         Width           =   11175
         Begin VB.TextBox txt_level_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1620
            Width           =   5145
         End
         Begin VB.TextBox txt_level_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1230
            Width           =   3795
         End
         Begin VB.TextBox txt_level_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   10
            TabIndex        =   4
            Top             =   810
            Width           =   1215
         End
         Begin VB.Label tlabel26 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   1950
            TabIndex        =   9
            Top             =   1650
            Width           =   1095
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   1950
            TabIndex        =   8
            Top             =   840
            Width           =   510
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   1950
            TabIndex        =   7
            Top             =   1260
            Width           =   525
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
         Height          =   3915
         Left            =   -74760
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
         TabIndex        =   46
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
         TabIndex        =   47
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
         TabIndex        =   55
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
      Begin TrueOleDBList60.TDBCombo TDBCombo_general 
         Height          =   375
         Left            =   1230
         OleObjectBlob   =   "frm_mst_allowance_range.frx":14443
         TabIndex        =   69
         Top             =   600
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   240
         TabIndex        =   71
         Top             =   5490
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew_Gen 
            Height          =   705
            Left            =   540
            TabIndex        =   72
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
            MICON           =   "frm_mst_allowance_range.frx":163A9
            PICN            =   "frm_mst_allowance_range.frx":163C5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_Gen 
            Height          =   705
            Left            =   1560
            TabIndex        =   73
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
            MICON           =   "frm_mst_allowance_range.frx":17457
            PICN            =   "frm_mst_allowance_range.frx":17473
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_Gen 
            Height          =   705
            Left            =   2580
            TabIndex        =   74
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
            MICON           =   "frm_mst_allowance_range.frx":18505
            PICN            =   "frm_mst_allowance_range.frx":18521
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_Gen 
            Height          =   705
            Left            =   3600
            TabIndex        =   75
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
            MICON           =   "frm_mst_allowance_range.frx":195B3
            PICN            =   "frm_mst_allowance_range.frx":195CF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_Gen 
            Height          =   705
            Left            =   4620
            TabIndex        =   76
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
            MICON           =   "frm_mst_allowance_range.frx":1A661
            PICN            =   "frm_mst_allowance_range.frx":1A67D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton CmdNew_Master_Gen 
            Height          =   705
            Left            =   8850
            TabIndex        =   77
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
            MICON           =   "frm_mst_allowance_range.frx":1B70F
            PICN            =   "frm_mst_allowance_range.frx":1B72B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_All_Gen 
            Height          =   705
            Left            =   9870
            TabIndex        =   1
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
            MICON           =   "frm_mst_allowance_range.frx":1C7BD
            PICN            =   "frm_mst_allowance_range.frx":1C7D9
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
         Left            =   -74730
         TabIndex        =   56
         Top             =   5490
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew_Travel 
            Height          =   705
            Left            =   540
            TabIndex        =   57
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
            MICON           =   "frm_mst_allowance_range.frx":1D86B
            PICN            =   "frm_mst_allowance_range.frx":1D887
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_Travel 
            Height          =   705
            Left            =   1560
            TabIndex        =   58
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
            MICON           =   "frm_mst_allowance_range.frx":1E919
            PICN            =   "frm_mst_allowance_range.frx":1E935
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_Travel 
            Height          =   705
            Left            =   2580
            TabIndex        =   59
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
            MICON           =   "frm_mst_allowance_range.frx":1F9C7
            PICN            =   "frm_mst_allowance_range.frx":1F9E3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_Travel 
            Height          =   705
            Left            =   3600
            TabIndex        =   60
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
            MICON           =   "frm_mst_allowance_range.frx":20A75
            PICN            =   "frm_mst_allowance_range.frx":20A91
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_Travel 
            Height          =   705
            Left            =   4620
            TabIndex        =   61
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
            MICON           =   "frm_mst_allowance_range.frx":21B23
            PICN            =   "frm_mst_allowance_range.frx":21B3F
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
      Begin VB.Frame fra_entry_general 
         Height          =   1815
         Left            =   240
         TabIndex        =   62
         Top             =   3570
         Visible         =   0   'False
         Width           =   11295
         Begin VB.TextBox txt_name_general 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   64
            Top             =   960
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
            TabIndex        =   65
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_general_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   10
            TabIndex        =   63
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   3480
            TabIndex        =   67
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   3795
            TabIndex        =   66
            Top             =   960
            Width           =   510
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Gen 
         Height          =   4335
         Left            =   240
         TabIndex        =   91
         Top             =   1020
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7646
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO"
         Columns(0).DataField=   "id_number"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "FROM (YEAR)"
         Columns(1).DataField=   "allow_under"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "TO (YEAR)"
         Columns(2).DataField=   "allow_upper"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   4
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "UP"
         Columns(3).DataField=   "flag_over"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "COMPENSATION"
         Columns(4).DataField=   "allow_compensation"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1270"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1191"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3043"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2963"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2831"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2752"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1429"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1349"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2355"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2275"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=8017"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=7938"
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
         Caption         =   "LIST OF ALLOWANCE RANGE GENERAL"
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
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=2"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Travel 
         Height          =   4875
         Left            =   -74730
         TabIndex        =   98
         Top             =   510
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   8599
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "TITLE CODE"
         Columns(0).DataField=   "title_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "JOB TITLE"
         Columns(1).DataField=   "title_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "ALLOWANCE (USD)"
         Columns(2).DataField=   "allowance_value"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2990"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2910"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=7064"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6985"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=3096"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3016"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=8678"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=8599"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=512"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
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
         Caption         =   "LIST OF TRAVELLING ALLOWANCE"
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
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=0"
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
         Caption         =   "TYPE"
         Height          =   195
         Left            =   240
         TabIndex        =   70
         Top             =   630
         Width           =   360
      End
   End
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   10680
      TabIndex        =   78
      Top             =   8100
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
      MICON           =   "frm_mst_allowance_range.frx":22BD1
      PICN            =   "frm_mst_allowance_range.frx":22BED
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER ALLOWANCE RANGE"
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
      Width           =   5625
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_allowance_range.frx":23C7F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11970
   End
End
Attribute VB_Name = "frm_mst_allowance_range"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAllow As New ADODB.Recordset
Dim rsAllow_Detail As New ADODB.Recordset
Dim rsAllow_Travel As New ADODB.Recordset

Dim rsTitle As New ADODB.Recordset
Dim vIdNumber As Integer

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim v_value_under, v_value_upper, v_value_percent, v_value As Double

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
    check_validate_exist_new = False
    
'    If SSTab1.Tab = 0 Then
'        SQL = "select count(id_number) as rec_count from m_allow_gen_detail where id_number = " _
'                & "'" & vIdNumber & "' AND allow_code = '" & TDBCombo_general.Text & "'"
'    Else
'        SQL = "select count(title_code) as rec_count from m_allow_travel where title_code = " _
'                & "'" & TDBCombo_title.Text & "'"
'    End If
'
'    rs.Open SQL, CnG, adOpenStatic, adLockReadOnly
'
'    If rs.Fields("rec_count").Value > 0 Then
'        check_validate_exist_new = True
'        Exit Function
'    End If
End Function

Private Sub check_invalid()
    MsgBox "Data found!", vbCritical, headerMSG
    If SSTab1.Tab = 1 Then
        TDBCombo_title.Text = ""
        txt_title_name.Text = ""
        If TDBCombo_title.Enabled = True Then TDBCombo_title.SetFocus
'    Else
'        txt_title_code = ""
'        If txt_title_code.Enabled = True Then txt_title_code.SetFocus
    End If
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
        
    If SSTab1.Tab = 1 Then
        If Not TDBCombo_title.Text = rsAllow_Travel.Fields("title_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
'    Else
'        If Not txt_title_code = rsAllow_Travel.Fields("title_code").Value And _
'        check_validate_exist_new Then
'            check_validate_exist_edit = True
'            Exit Function
'        End If
    End If
    
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True

    If SSTab1.Tab = 0 Then
        If Trim(txtFrom_Gen) = "" Then
            MsgBox "From Value is empty!", vbOKOnly + vbInformation, headerMSG
            txtFrom_Gen.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If chkUp.Value = 0 Then
            If Trim(txtTo_Gen.Text) = "" Then
                MsgBox "To Value is empty!", vbOKOnly + vbInformation, headerMSG
                txtTo_Gen.SetFocus
                check_validate_new = False
                Exit Function
            End If
        End If
        
        If Trim(txtCompensation_Gen.Text) = "" Then
            MsgBox "Compensation Value is empty!", vbOKOnly + vbInformation, headerMSG
            txtCompensation_Gen.SetFocus
            check_validate_new = False
            Exit Function
        End If
    Else
        If Trim(TDBCombo_title.Text) = "" Then
            MsgBox "Title Code is empty!", vbOKOnly + vbInformation, headerMSG
            TDBCombo_title.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If Trim(txt_allow_travel.Text) = "" Then
            MsgBox "Allowance value is empty!", vbOKOnly + vbInformation, headerMSG
            txt_allow_travel.SetFocus
            check_validate_new = False
            Exit Function
        End If
    End If
End Function

Private Sub cancel_data()
    int_mode = 0
    Call load_mode
    
    CmdNew_Master_Gen.Caption = "&New"
    cmdCancel_Gen.Caption = "&Cancel Dtl"
End Sub

Private Sub delete_all_data()
Dim i As Integer

On Error GoTo Err
    If SSTab1.Tab = 0 Then
        i = MsgBox("Are you sure want to delete data '" _
            & txt_general.Text & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_allow_gen_detail where " _
                & "allow_code = '" & TDBCombo_general.Text & "'"
        CnG.Execute "delete from m_allow_gen where allow_code = " _
                & "'" & TDBCombo_general.Text & "'"
        
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        SQL = "Update temp_sal_proses set salary_proses = 0"
        CnG.Execute SQL
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        CnG.CommitTrans
        
        Call load_data_master
        Call load_data_detail
        int_mode = 0
        Call load_mode
        
        TDBCombo_general.Text = ""
        txt_name_general = ""
        Set TDBGrid_Gen.DataSource = Nothing
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub delete_data()
Dim i As Integer
'On Error GoTo Err
    If SSTab1.Tab = 0 Then
        If Not (TDBGrid_Gen.ApproxCount > 0 And TDBGrid_Gen.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Are you sure want to delete data '" _
            & TDBGrid_Gen.Columns("allow_compensation").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_allow_gen_detail where id_number = " _
                & "'" & TDBGrid_Gen.Columns("id_number").Value & "' AND allow_code = '" & TDBCombo_general.Text & "'"
        
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        SQL = "Update temp_sal_proses set salary_proses = 0"
        CnG.Execute SQL
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        CnG.CommitTrans
        
        Call load_data_detail
        int_mode = 0
        Call load_mode
    Else
        If Not (TDBGrid_Travel.ApproxCount > 0 And TDBGrid_Travel.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Are you sure want to delete data '" _
            & TDBGrid_Travel.Columns("title_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_allow_travel where title_code = " _
                & "'" & TDBGrid_Travel.Columns("title_code").Value & "'"
        
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
        If Not (TDBGrid_Gen.ApproxCount > 0 And TDBGrid_Gen.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsAllow_Detail
            vIdNumber = .Fields("id_number").Value
            txtFrom_Gen.Text = FormatNumber(.Fields("allow_under").Value)
            txtTo_Gen.Text = FormatNumber(.Fields("allow_upper").Value)
            chkUp.Value = .Fields("flag_over").Value
            txtCompensation_Gen.Text = FormatNumber(.Fields("allow_compensation").Value)
            txtDescription_Gen.Text = .Fields("description").Value
        End With
        
        v_value_under = txtFrom_Gen.Text
        v_value_upper = txtTo_Gen.Text
        v_value_percent = txtCompensation_Gen.Text
    Else
        If Not (TDBGrid_Travel.ApproxCount > 0 And TDBGrid_Travel.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsAllow_Travel
            TDBCombo_title.Text = .Fields("title_code").Value
            txt_title_name.Text = .Fields("title_name").Value
            txt_allow_travel.Text = .Fields("allowance_value").Value
            txt_description_travel.Text = .Fields("description").Value
        End With
        
        v_value = txt_allow_travel.Text
    End If
End Sub

Private Sub edit_data()
    int_mode = 2
    Call load_mode
End Sub

Private Sub chkUp_Click()
    If chkUp.Value = 1 Then
        txtTo_Gen.Enabled = False
        txtTo_Gen.Text = ""
    Else
        txtTo_Gen.Enabled = True
    End If
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
    If SSTab1.Tab = 0 Then
        If CmdNew_Master_Gen.Caption = "&New" Then
            CmdNew_Master_Gen.Caption = "&Save"
            cmdCancel_Gen.Caption = "&Cancel"
            Call set_buttons_enable(False, False, False, False, True, False, False)
            fra_entry_general.Visible = True
            fra_entry_dtl_general.Visible = False
            
            txt_general_code.Text = ""
            txt_name_general.Text = ""
            txt_general_code.SetFocus
            
            CmdNew_Master_Gen.Enabled = True
        Else
            SQL = "INSERT INTO m_allow_gen(allow_code,allow_name) " _
                    & "VALUES ('" & txt_general_code & "','" & txt_name_general & "')"
            CnG.Execute SQL
            
            Call set_buttons_enable(True, False, True, True, False, True, True)
            CmdNew_Master_Gen.Caption = "&New"
            cmdCancel_Gen.Caption = "&Cancel Dtl"
                
            fra_entry_general.Visible = False
            
            txt_general_code.Text = ""
            txt_name_general.Text = ""
        
            Call load_data_master
            txt_general.Text = ""
        End If
    End If
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub insert_new_data()
Dim vAllowFrom As Long
Dim vAlloTo As Long

On Error GoTo Err
    CnG.BeginTrans
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    SQL = "Update temp_sal_proses set salary_proses = 0"
    CnG.Execute SQL
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    If SSTab1.Tab = 0 Then
        
        vAllowFrom = Val(Replace(txtFrom_Gen.Text, ",", ""))
        vAlloTo = Val(Replace(txtTo_Gen.Text, ",", ""))
        
        If vAllowFrom > vAlloTo And chkUp = 0 Then
            MsgBox "To Value is smaller than From Value!", vbExclamation, headerMSG
            txtTo_Gen.SetFocus
            Exit Sub
        End If
        
        If Not Trim(txtTo_Gen.Text) = "" Then
            txtTo_Gen.Text = FormatNumber(DropAllComma(txtTo_Gen.Text))
        End If
    
        SQL = "SELECT MAX(id_number) id_number FROM m_allow_gen_detail " & _
                "WHERE allow_code = '" & TDBCombo_general.Text & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            vIdNumber = IIf(IsNull(rs!id_number), 0, rs!id_number) + 1
        End If
        rs.Close
        
        SQL = "INSERT INTO m_allow_gen_detail (allow_code,id_number,allow_under,allow_upper," & _
                "flag_over,allow_compensation,description) " & _
              "VALUES( " & _
                "'" & Trim(TDBCombo_general.Text) & "','" & vIdNumber & "','" & Val(DropAllComma(txtFrom_Gen.Text)) & "','" & Val(DropAllComma(txtTo_Gen.Text)) & "'," & _
                "'" & chkUp.Value & "','" & Val(DropAllComma(txtCompensation_Gen.Text)) & "'," & _
                "'" & Trim(txtDescription_Gen.Text) & "')"
        CnG.Execute SQL
    Else
        SQL = "INSERT INTO m_allow_travel (title_code,allowance_value,description) " & _
              "VALUES( " & _
                "'" & Trim(TDBCombo_title.Text) & "','" & Val(DropAllComma(txt_allow_travel.Text)) & "'," & _
                "'" & Trim(txt_description_travel.Text) & "')"
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
    If SSTab1.Tab = 0 Then
        If v_value_under <> txtFrom_Gen.Text Or v_value_upper <> txtTo_Gen.Text Or v_value_percent <> txtCompensation_Gen.Text Then
            SQL = "Update temp_sal_proses set salary_proses = 0"
            CnG.Execute SQL
        End If
        
        SQL = "UPDATE m_allow_gen_detail SET allow_code = '" & Trim(TDBCombo_general.Text) & "'," & _
                "allow_under = '" & Val(DropAllComma(txtFrom_Gen.Text)) & "'," & _
                "allow_upper = '" & Val(DropAllComma(txtTo_Gen.Text)) & "'," & _
                "flag_over = '" & IIf(chkUp, 1, 0) & "'," & _
                "allow_compensation = '" & Val(DropAllComma(txtCompensation_Gen.Text)) & "'," & _
                "description = '" & Trim(txtDescription_Gen.Text) & "' " & _
              "WHERE allow_code = '" & TDBCombo_general.Text & "' " & _
                "AND id_number = '" & vIdNumber & "'"
        CnG.Execute SQL
    Else
        If v_value <> txt_allow_travel.Text Then
            SQL = "Update temp_sal_proses set salary_proses = 0"
            CnG.Execute SQL
        End If
        
        SQL = "UPDATE m_allow_travel SET allowance_value = '" & Val(DropAllComma(txt_allow_travel.Text)) & "'," & _
                "description = '" & Trim(txt_description_travel.Text) & "' " & _
              "WHERE title_code = '" & Trim(TDBCombo_title.Text) & "'"
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
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
    If SSTab1.Tab = 0 Then
        cmdNew_Gen.Enabled = a And blnUser_Add
        cmdSave_Gen.Enabled = b
        cmdEdit_Gen.Enabled = c And blnUser_Edit
        cmdDelete_Gen.Enabled = d And blnUser_Delete
        cmdCancel_Gen.Enabled = e
        
        CmdNew_Master_Gen.Enabled = a And blnUser_Add
        cmdDelete_All_Gen.Enabled = d And blnUser_Delete
    Else
        cmdNew_Travel.Enabled = a And blnUser_Add
        cmdSave_Travel.Enabled = b
        cmdEdit_Travel.Enabled = c And blnUser_Edit
        cmdDelete_Travel.Enabled = d And blnUser_Delete
        cmdCancel_Travel.Enabled = e
    End If
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If SSTab1.Tab = 0 Then
                If Not LCase(Ctr.name) = "txt_general" Then Ctr.Text = ""
            Else
                If Not LCase(Ctr.name) = "txt_pph21" Then Ctr.Text = ""
            End If
        ElseIf TypeOf Ctr Is TDBCombo Then
            If SSTab1.Tab = 0 Then
                If Not LCase(Ctr.name) = "tdbcombo_general" Then Ctr.Text = ""
            Else
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
        If SSTab1.Tab = 0 Then
            If Trim(TDBCombo_general) = "" Then
            MsgBox "Type is not selected!", vbOKOnly + vbInformation, headerMSG
            TDBCombo_general.SetFocus
            
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
    
        fra_entry_dtl_general.Visible = True
        cmdCancel_Gen.Caption = "&Cancel Dtl"
        TDBGrid_Gen.Enabled = False
        Call set_new_data
        
        txtFrom_Gen.SetFocus
    Else
        fra_entry_Travel.Visible = True
        TDBCombo_title.Enabled = True
        TDBGrid_Travel.Enabled = False
        Call set_new_data
        
        If TDBCombo_title.Enabled = True Then
            TDBCombo_title.SetFocus
        End If
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    If SSTab1.Tab = 0 Then
        fra_entry_general.Visible = False
        fra_entry_dtl_general.Visible = False
        TDBGrid_Gen.Enabled = True
    Else
        fra_entry_Travel.Visible = False
        TDBGrid_Travel.Enabled = True
    End If

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    
    If vSetData = 0 Then
        int_mode = 0
        Call load_mode
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        fra_entry_dtl_general.Visible = True
        TDBGrid_Gen.Enabled = False
    Else
        TDBCombo_title.Enabled = False
        fra_entry_Travel.Visible = True
        TDBGrid_Travel.Enabled = False
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
    oClause = ""
    If SSTab1.Tab = 0 Then
        Call load_data_master
    Else
        Call load_data_detail
    End If
        
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    
    SSTab1.Tab = 0
End Sub

Private Sub clear_filter()
    If SSTab1.Tab = 0 Then
        For Each Col In TDBGrid_Gen.Columns
            Col.FilterText = ""
        Next Col
        rsAllow_Detail.Filter = adFilterNone
    Else
        For Each Col In TDBGrid_Travel.Columns
            Col.FilterText = ""
        Next Col
        rsAllow_Travel.Filter = adFilterNone
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
    If SSTab1.Tab = 0 Then
        Set Cols = TDBGrid_Gen.Columns
        i = TDBGrid_Gen.Col
        TDBGrid_Gen.HoldFields
        
        rsAllow_Detail.Filter = getFilter()
        TDBGrid_Gen.Col = i
        TDBGrid_Gen.EditActive = True
        
        TDBGrid_Gen.SelStart = Len(TDBGrid_Gen.Columns(i).FilterText)
        If TDBGrid_Gen.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Gen.Col = i
        End If
    Else
        Set Cols = TDBGrid_Travel.Columns
        i = TDBGrid_Travel.Col
        TDBGrid_Travel.HoldFields
        
        rsAllow_Travel.Filter = getFilter()
        TDBGrid_Travel.Col = i
        TDBGrid_Travel.EditActive = True
        
        TDBGrid_Travel.SelStart = Len(TDBGrid_Travel.Columns(i).FilterText)
        If TDBGrid_Travel.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Travel.Col = i
        End If
    End If
    
    Exit Sub
    
Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data_title()
    If rsTitle.State Then rsTitle.Close
    SQL = "select * from m_title order by title_code"
    rsTitle.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_title.RowSource = rsTitle
End Sub

Private Sub load_data_detail()
    If SSTab1.Tab = 0 Then
        If rsAllow_Detail.State Then rsAllow_Detail.Close
        SQL = "select * from m_allow_gen_detail " & _
                "where allow_code = '" & TDBCombo_general.Text & "' " & oClause
'        SQL = "SELECT allow_code,id_number,CONCAT('> ',allow_under) allow_under," & _
'                "CONCAT('< ',allow_upper) allow_upper, flag_over," & _
'                "CONCAT(allow_compensation, ' bulan gaji pokok') allow_compensation," & _
'                "Description " & _
'              "from m_allow_gen_detail " & _
'              "WHERE allow_code = '" & TDBCombo_general.Text & "' " & oClause
        rsAllow_Detail.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Gen.DataSource = rsAllow_Detail
    Else
        If rsAllow_Travel.State Then rsAllow_Travel.Close
        SQL = "select a.*,b.title_name from m_allow_travel a join m_title b on a.title_code = b.title_code " & oClause
        rsAllow_Travel.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Travel.DataSource = rsAllow_Travel
    End If
End Sub

Private Sub load_data_master()
    If SSTab1.Tab = 0 Then
        If rsAllow.State Then rsAllow.Close
        SQL = "select * from m_allow_gen order by allow_code"
        rsAllow.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBCombo_general.RowSource = rsAllow
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_allowance_range = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    int_mode = 0
    Call load_mode
    
    oClause = ""
    If SSTab1.Tab = 0 Then
        TDBGrid_Travel.DataSource = Nothing
        Call load_data_master
    Else
        TDBGrid_Gen.DataSource = Nothing
        Call load_data_detail
        Call load_data_title
    End If
End Sub

Private Sub TDBCombo_general_ItemChange()
    If TDBCombo_general.ApproxCount > 0 Then
        TDBCombo_general.Text = TDBCombo_general.Columns("allow_code").Value
        txt_general = TDBCombo_general.Columns("allow_name").Value
        
        Call load_data_detail
    End If
End Sub

Private Sub TDBCombo_title_ItemChange()
    If TDBCombo_title.ApproxCount > 0 Then
        TDBCombo_title.Text = TDBCombo_title.Columns("title_code").Value
        txt_title_name.Text = TDBCombo_title.Columns("title_name").Value
    End If
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

Private Sub CmdNew_Master_Gen_Click()
    Call new_data_master
End Sub

Private Sub cmdDelete_All_Gen_Click()
    Call delete_all_data
End Sub


Private Sub cmdNew_Travel_Click()
    Call new_data
End Sub

Private Sub cmdSave_Travel_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Travel_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Travel_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Travel_Click()
    Call cancel_data
End Sub


Private Sub TDBGrid_Travel_FilterChange()
    Call filter_change
End Sub

Private Sub TDBGrid_Gen_FilterChange()
    Call filter_change
End Sub


Private Sub txt_general_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_name_general_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFrom_Gen_Validate(Cancel As Boolean)
    If Not Trim(txtFrom_Gen.Text) = "" Then
        txtFrom_Gen.Text = FormatNumber(DropAllComma(txtFrom_Gen.Text))
    End If
End Sub

Private Sub txtTo_Gen_Validate(Cancel As Boolean)
    If Not Trim(txtTo_Gen.Text) = "" Then
        txtTo_Gen.Text = FormatNumber(DropAllComma(txtTo_Gen.Text))
    End If
End Sub

'Private Sub txtTo_Gen_Validate(Cancel As Boolean)
'Dim vAllowFrom As Long
'Dim vAlloTo As Long
'
'    vAllowFrom = Val(Replace(txtFrom_Gen.Text, ",", ""))
'    vAlloTo = Val(Replace(txtTo_Gen.Text, ",", ""))
'
'    If vAllowFrom > vAlloTo And chkUp = 0 Then
'        MsgBox "To Value is smaller than From Value!", vbExclamation, headerMSG
'        txtTo_Gen.SetFocus
'        Exit Sub
'    End If
'
'    If Not Trim(txtTo_Gen.Text) = "" Then
'        txtTo_Gen.Text = FormatNumber(DropAllComma(txtTo_Gen.Text))
'    End If
'End Sub

Private Sub txtCompensation_Gen_Validate(Cancel As Boolean)
    If Not Trim(txtCompensation_Gen.Text) = "" Then
        txtCompensation_Gen.Text = FormatNumber(DropAllComma(txtCompensation_Gen.Text))
    End If
End Sub


Private Sub txt_allow_travel_Validate(Cancel As Boolean)
    If Not Trim(txt_allow_travel.Text) = "" Then
        txt_allow_travel.Text = FormatNumber(DropAllComma(txt_allow_travel.Text))
    End If
End Sub

Private Sub TDBGrid_Gen_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Gen.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Gen.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Gen.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Gen.Columns(ColIndex).DataField
    Call load_data_detail

End Sub

Private Sub TDBGrid_Travel_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Travel.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Travel.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Travel.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Travel.Columns(ColIndex).DataField
    Call load_data_detail

End Sub
