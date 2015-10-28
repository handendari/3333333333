VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_preferensi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER PREFERENCE"
   ClientHeight    =   7995
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
   Icon            =   "frm_mst_preferensi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   90
      Top             =   7440
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   11139
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "GENERAL"
      TabPicture(0)   =   "frm_mst_preferensi.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(1)=   "Frame4"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "OVERTIME"
      TabPicture(1)   =   "frm_mst_preferensi.frx":05A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "PRESENCE"
      TabPicture(2)   =   "frm_mst_preferensi.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TDBGrid_Presence"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "fra_entry_jhk"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "UMK"
      TabPicture(3)   =   "frm_mst_preferensi.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TDBGrid_UMK"
      Tab(3).Control(1)=   "frmTombol"
      Tab(3).Control(2)=   "fra_entry"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "LEAVE"
      TabPicture(4)   =   "frm_mst_preferensi.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TDBGrid_Leave"
      Tab(4).Control(1)=   "fra_entry_leave"
      Tab(4).Control(2)=   "Frame3"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "PENALTY"
      TabPicture(5)   =   "frm_mst_preferensi.frx":0616
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "SSTab2"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "ALLOWANCE"
      TabPicture(6)   =   "frm_mst_preferensi.frx":0632
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame15"
      Tab(6).Control(1)=   "Frame16"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "PERIODE"
      TabPicture(7)   =   "frm_mst_preferensi.frx":064E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame21"
      Tab(7).Control(1)=   "Frame20"
      Tab(7).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   4335
         Left            =   1890
         TabIndex        =   64
         Top             =   480
         Width           =   7965
         Begin VB.TextBox txt_change_status 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5700
            TabIndex        =   81
            Top             =   1590
            Width           =   675
         End
         Begin VB.CheckBox chk_approve 
            Caption         =   "AUTO APPROVED"
            Height          =   255
            Left            =   6090
            TabIndex        =   72
            Top             =   210
            Width           =   1575
         End
         Begin VB.TextBox txt_on_call_allowance 
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
            Height          =   285
            Left            =   2370
            MaxLength       =   10
            TabIndex        =   80
            Top             =   1590
            Width           =   1275
         End
         Begin VB.TextBox txtMAXOt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   2370
            TabIndex        =   76
            Top             =   870
            Width           =   675
         End
         Begin VB.CheckBox chk_transport 
            Caption         =   "YES"
            Height          =   255
            Left            =   6090
            TabIndex        =   77
            Top             =   870
            Width           =   645
         End
         Begin VB.TextBox txtAutoOt_wd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2370
            TabIndex        =   78
            Top             =   1260
            Width           =   675
         End
         Begin VB.TextBox txtAutoOt_Sat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5700
            TabIndex        =   79
            Top             =   1260
            Width           =   675
         End
         Begin VB.CheckBox chk_meal 
            Caption         =   "YES"
            Height          =   255
            Left            =   6090
            TabIndex        =   75
            Top             =   540
            Width           =   645
         End
         Begin VB.TextBox txtMAXOt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2370
            TabIndex        =   74
            Top             =   540
            Width           =   675
         End
         Begin VB.Frame fraSystem 
            Height          =   2385
            Left            =   1170
            TabIndex        =   71
            Top             =   1860
            Visible         =   0   'False
            Width           =   5535
            Begin VB.TextBox txt_ot_name3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   2535
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   97
               Top             =   1950
               Width           =   2835
            End
            Begin VB.TextBox txt_ot_name2 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   2535
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   95
               Top             =   1590
               Width           =   2835
            End
            Begin VB.TextBox txt_ot_name1 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   2535
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   93
               Top             =   1140
               Width           =   2835
            End
            Begin VB.Frame Frame7 
               BorderStyle     =   0  'None
               Height          =   645
               Left            =   510
               TabIndex        =   91
               Top             =   450
               Width           =   3975
               Begin VB.TextBox txtOTStart 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   390
                  TabIndex        =   84
                  Top             =   360
                  Width           =   495
               End
               Begin VB.OptionButton optDelay 
                  Height          =   195
                  Left            =   120
                  TabIndex        =   83
                  Top             =   420
                  Width           =   285
               End
               Begin VB.OptionButton optAfter 
                  Caption         =   "After Work Day"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   82
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   1425
               End
               Begin VB.Label Label20 
                  Caption         =   "Hours After Work Day"
                  Height          =   225
                  Left            =   960
                  TabIndex        =   92
                  Top             =   390
                  Width           =   1635
               End
            End
            Begin TrueOleDBList60.TDBCombo TDBCombo_ot1 
               Height          =   375
               Left            =   1275
               OleObjectBlob   =   "frm_mst_preferensi.frx":066A
               TabIndex        =   85
               Top             =   1140
               Width           =   1245
            End
            Begin TrueOleDBList60.TDBCombo TDBCombo_ot2 
               Height          =   375
               Left            =   1275
               OleObjectBlob   =   "frm_mst_preferensi.frx":25B4
               TabIndex        =   86
               Top             =   1590
               Width           =   1245
            End
            Begin TrueOleDBList60.TDBCombo TDBCombo_ot3 
               Height          =   375
               Left            =   1275
               OleObjectBlob   =   "frm_mst_preferensi.frx":44FE
               TabIndex        =   87
               Top             =   1950
               Width           =   1245
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "OT SUNDAY*"
               Height          =   195
               Left            =   300
               TabIndex        =   98
               Top             =   1980
               Width           =   945
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "OT SATURDAY*"
               Height          =   195
               Left            =   105
               TabIndex        =   96
               Top             =   1590
               Width           =   1140
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "OT TYPE*"
               Height          =   195
               Left            =   510
               TabIndex        =   94
               Top             =   1170
               Width           =   705
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "OVERTIME CALCULATION START"
               Height          =   225
               Left            =   120
               TabIndex        =   90
               Top             =   270
               Width           =   2505
            End
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3570
            TabIndex        =   65
            Top             =   150
            Width           =   2415
            Begin VB.OptionButton optSPL 
               Caption         =   "SPL"
               Height          =   225
               Left            =   210
               TabIndex        =   67
               Top             =   90
               Value           =   -1  'True
               Width           =   645
            End
            Begin VB.OptionButton optSystem 
               Caption         =   "SYSTEM"
               Height          =   225
               Left            =   1050
               TabIndex        =   66
               Top             =   90
               Width           =   915
            End
         End
         Begin VB.Label Label69 
            Caption         =   "Hours"
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
            Left            =   6450
            TabIndex        =   226
            Top             =   1650
            Width           =   555
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CHANGE PRESENT TO OT"
            Height          =   195
            Index           =   1
            Left            =   3825
            TabIndex        =   225
            Top             =   1650
            Width           =   1830
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ON CALL ALLOWANCE"
            Height          =   195
            Index           =   0
            Left            =   705
            TabIndex        =   221
            Top             =   1650
            Width           =   1605
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            Caption         =   "FOR OT MIN"
            Height          =   225
            Left            =   510
            TabIndex        =   209
            Top             =   900
            Width           =   1755
         End
         Begin VB.Label Label61 
            Caption         =   "Hours"
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
            Left            =   3150
            TabIndex        =   208
            Top             =   900
            Width           =   555
         End
         Begin VB.Label Label60 
            Alignment       =   1  'Right Justify
            Caption         =   "GET TRANSPORT ALLOWANCE"
            Height          =   225
            Left            =   3750
            TabIndex        =   207
            Top             =   900
            Width           =   2235
         End
         Begin VB.Label Label27 
            Caption         =   "Hours"
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
            Left            =   3150
            TabIndex        =   113
            Top             =   1290
            Width           =   555
         End
         Begin VB.Label Label26 
            Caption         =   "Hours"
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
            Left            =   6450
            TabIndex        =   112
            Top             =   1290
            Width           =   555
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "AUTO OT WORK DAY"
            Height          =   195
            Left            =   780
            TabIndex        =   111
            Top             =   1320
            Width           =   1530
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "AUTO OT SATURDAY"
            Height          =   195
            Left            =   4110
            TabIndex        =   110
            Top             =   1320
            Width           =   1515
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "GET MEAL ALLOWANCE"
            Height          =   225
            Left            =   3750
            TabIndex        =   89
            Top             =   570
            Width           =   2235
         End
         Begin VB.Label Label17 
            Caption         =   "Hours"
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
            Left            =   3150
            TabIndex        =   88
            Top             =   570
            Width           =   555
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "FOR OT MIN"
            Height          =   225
            Left            =   510
            TabIndex        =   73
            Top             =   570
            Width           =   1755
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "OVERTIME METHOD"
            Height          =   225
            Left            =   1710
            TabIndex        =   68
            Top             =   240
            Width           =   1665
         End
      End
      Begin VB.Frame Frame21 
         Height          =   3255
         Left            =   -73350
         TabIndex        =   212
         Top             =   750
         Width           =   7965
         Begin VB.OptionButton optPeriode 
            Caption         =   "PERIODE"
            Height          =   255
            Index           =   1
            Left            =   3630
            TabIndex        =   220
            Top             =   1410
            Width           =   1035
         End
         Begin VB.OptionButton optPeriode 
            Caption         =   "MONTHLY"
            Height          =   255
            Index           =   0
            Left            =   3630
            TabIndex        =   219
            Top             =   1140
            Width           =   1035
         End
         Begin VB.Frame fraPeriode 
            Height          =   915
            Left            =   3630
            TabIndex        =   214
            Top             =   1620
            Width           =   2235
            Begin VB.TextBox txt_day_end 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1410
               TabIndex        =   218
               Top             =   480
               Width           =   345
            End
            Begin VB.TextBox txt_day_start 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1410
               TabIndex        =   216
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label65 
               Caption         =   "End Date :"
               Height          =   195
               Left            =   450
               TabIndex        =   217
               Top             =   510
               Width           =   945
            End
            Begin VB.Label Label64 
               Caption         =   "Start Date :"
               Height          =   195
               Left            =   450
               TabIndex        =   215
               Top             =   210
               Width           =   945
            End
         End
         Begin VB.Label Label63 
            Caption         =   "PERIOD SALARY :"
            Height          =   285
            Left            =   2220
            TabIndex        =   213
            Top             =   1140
            Width           =   1485
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -73350
         TabIndex        =   210
         Top             =   4110
         Width           =   7965
         Begin prj_tpc.vbButton cmdSave_Periode 
            Height          =   705
            Left            =   5820
            TabIndex        =   211
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":6448
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":6474
            picn            =   "frm_mst_preferensi.frx":6492
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
      Begin VB.Frame Frame16 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -73260
         TabIndex        =   184
         Top             =   4230
         Width           =   7965
         Begin prj_tpc.vbButton cmdSave_Allow 
            Height          =   705
            Left            =   5820
            TabIndex        =   194
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":7526
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":7552
            picn            =   "frm_mst_preferensi.frx":7570
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
      Begin VB.Frame Frame15 
         Height          =   3255
         Left            =   -73260
         TabIndex        =   183
         Top             =   870
         Width           =   7965
         Begin VB.TextBox txt_iuran_koperasi 
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
            Left            =   5220
            MaxLength       =   10
            TabIndex        =   192
            Top             =   2340
            Width           =   2325
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "APPLY TO ALL"
            Height          =   315
            Left            =   1380
            TabIndex        =   206
            Top             =   2760
            Width           =   1545
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
            Left            =   5220
            MaxLength       =   10
            TabIndex        =   190
            Top             =   1620
            Width           =   2325
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
            Left            =   5220
            MaxLength       =   10
            TabIndex        =   191
            Top             =   1980
            Width           =   2325
         End
         Begin VB.Frame Frame19 
            Height          =   525
            Left            =   5205
            TabIndex        =   202
            Top             =   720
            Width           =   2325
            Begin VB.OptionButton opt_daily_transport 
               Caption         =   "DAILY"
               Height          =   225
               Left            =   90
               TabIndex        =   200
               Top             =   210
               Value           =   -1  'True
               Width           =   1005
            End
            Begin VB.OptionButton opt_monthly_transport 
               Caption         =   "MONTHLY"
               Height          =   225
               Left            =   1110
               TabIndex        =   201
               Top             =   210
               Width           =   1125
            End
         End
         Begin VB.TextBox txt_transport_allowance 
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
            Left            =   5205
            MaxLength       =   10
            TabIndex        =   189
            Top             =   1260
            Width           =   2325
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
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   187
            Top             =   1260
            Width           =   2325
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
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   188
            Top             =   2130
            Width           =   2325
         End
         Begin VB.Frame Frame18 
            Height          =   525
            Left            =   1380
            TabIndex        =   186
            Top             =   720
            Width           =   2325
            Begin VB.OptionButton opt_daily_presence 
               Caption         =   "DAILY"
               Height          =   225
               Left            =   90
               TabIndex        =   193
               Top             =   210
               Width           =   1005
            End
            Begin VB.OptionButton opt_monthly_presence 
               Caption         =   "MONTHLY"
               Height          =   225
               Left            =   1110
               TabIndex        =   195
               Top             =   210
               Value           =   -1  'True
               Width           =   1125
            End
         End
         Begin VB.Frame Frame17 
            Height          =   525
            Left            =   1380
            TabIndex        =   185
            Top             =   1590
            Width           =   2325
            Begin VB.OptionButton opt_monthly_meal 
               Caption         =   "MONTHLY"
               Height          =   225
               Left            =   1110
               TabIndex        =   199
               Top             =   210
               Width           =   1125
            End
            Begin VB.OptionButton opt_daily_meal 
               Caption         =   "DAILY"
               Height          =   225
               Left            =   90
               TabIndex        =   197
               Top             =   210
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.Label Label67 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "IURAN KOPERASI"
            Height          =   195
            Left            =   3825
            TabIndex        =   222
            Top             =   2400
            Width           =   1275
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SHIFT 3"
            Height          =   195
            Left            =   3360
            TabIndex        =   205
            Top             =   2040
            Width           =   1740
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SHIFT 2"
            Height          =   195
            Left            =   3360
            TabIndex        =   204
            Top             =   1680
            Width           =   1740
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TRANSPORT"
            Height          =   195
            Left            =   4080
            TabIndex        =   203
            Top             =   1140
            Width           =   1005
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PRESENCE (%)"
            Height          =   195
            Left            =   195
            TabIndex        =   198
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MEAL"
            Height          =   195
            Left            =   825
            TabIndex        =   196
            Top             =   2010
            Width           =   435
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   5535
         Left            =   -74760
         TabIndex        =   114
         Top             =   510
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   9763
         _Version        =   393216
         Style           =   1
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "PRESENCE ALLOWANCE"
         TabPicture(0)   =   "frm_mst_preferensi.frx":8604
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "TDBGrid_PreAll"
         Tab(0).Control(1)=   "fra_entry_PreAll"
         Tab(0).Control(2)=   "Frame10"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "TIME LATE CONVERTION"
         TabPicture(1)   =   "frm_mst_preferensi.frx":8620
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TDBGrid_LateConvert"
         Tab(1).Control(1)=   "Frame9"
         Tab(1).Control(2)=   "fra_entry_LateConvert"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "GENERAL"
         TabPicture(2)   =   "frm_mst_preferensi.frx":863C
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "cmdSave_Gen"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame11"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Frame12"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Frame13"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Frame14"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).ControlCount=   5
         Begin VB.Frame Frame14 
            Caption         =   "ALPHA"
            Height          =   1005
            Left            =   2550
            TabIndex        =   168
            Top             =   3990
            Width           =   6765
            Begin VB.TextBox txtAlpha 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1170
               TabIndex        =   158
               Top             =   240
               Width           =   495
            End
            Begin VB.CheckBox chkAlpha 
               Caption         =   "YES"
               Height          =   195
               Left            =   4080
               TabIndex        =   159
               Top             =   630
               Width           =   765
            End
            Begin VB.Label Label49 
               Caption         =   "x ALPHA x BASIC SALARY / 30"
               Height          =   195
               Left            =   1740
               TabIndex        =   170
               Top             =   300
               Width           =   2295
            End
            Begin VB.Label Label48 
               Caption         =   "TAKE OFF PRESENCE ALLOWANCE"
               Height          =   225
               Left            =   1200
               TabIndex        =   169
               Top             =   630
               Width           =   2505
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "SICK TOLERANCE"
            Height          =   1005
            Left            =   2550
            TabIndex        =   165
            Top             =   2910
            Width           =   6765
            Begin VB.TextBox txtSick 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1170
               TabIndex        =   156
               Top             =   240
               Width           =   495
            End
            Begin VB.CheckBox chkSick 
               Caption         =   "YES"
               Height          =   195
               Left            =   4080
               TabIndex        =   157
               Top             =   630
               Width           =   765
            End
            Begin VB.Label Label47 
               Caption         =   "MAX ANNUALLY"
               Height          =   195
               Left            =   1740
               TabIndex        =   167
               Top             =   300
               Width           =   3375
            End
            Begin VB.Label Label46 
               Caption         =   "TAKE OFF PRESENCE ALLOWANCE"
               Height          =   225
               Left            =   1200
               TabIndex        =   166
               Top             =   630
               Width           =   2505
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "PRIVATE LEAVE (NOT APPROVED)"
            Height          =   1005
            Left            =   2550
            TabIndex        =   162
            Top             =   1830
            Width           =   6765
            Begin VB.TextBox txtPL_Not 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1170
               TabIndex        =   154
               Top             =   240
               Width           =   495
            End
            Begin VB.CheckBox chkPL_Not 
               Caption         =   "YES"
               Height          =   195
               Left            =   4080
               TabIndex        =   155
               Top             =   630
               Width           =   765
            End
            Begin VB.Label lblPL_Not 
               Caption         =   "..."
               Height          =   225
               Left            =   5280
               TabIndex        =   172
               Top             =   300
               Width           =   555
            End
            Begin VB.Label Label45 
               Caption         =   "x PERMISSION HOUR TOTAL x BASIC SALARY /"
               Height          =   195
               Left            =   1740
               TabIndex        =   164
               Top             =   300
               Width           =   3435
            End
            Begin VB.Label Label44 
               Caption         =   "TAKE OFF PRESENCE ALLOWANCE"
               Height          =   225
               Left            =   1200
               TabIndex        =   163
               Top             =   630
               Width           =   2505
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "PRIVATE LEAVE"
            Height          =   1005
            Left            =   2550
            TabIndex        =   151
            Top             =   750
            Width           =   6765
            Begin VB.CheckBox chkPL 
               Caption         =   "YES"
               Height          =   195
               Left            =   4080
               TabIndex        =   153
               Top             =   630
               Width           =   765
            End
            Begin VB.TextBox txtPL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1170
               TabIndex        =   152
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lblPL 
               Caption         =   "..."
               Height          =   225
               Left            =   5250
               TabIndex        =   171
               Top             =   300
               Width           =   555
            End
            Begin VB.Label Label43 
               Caption         =   "TAKE OFF PRESENCE ALLOWANCE"
               Height          =   225
               Left            =   1200
               TabIndex        =   161
               Top             =   630
               Width           =   2505
            End
            Begin VB.Label Label42 
               Caption         =   "x PERMISSION HOUR TOTAL x BASIC SALARY /"
               Height          =   195
               Left            =   1740
               TabIndex        =   160
               Top             =   300
               Width           =   3465
            End
         End
         Begin VB.Frame fra_entry_LateConvert 
            Height          =   2475
            Left            =   -74820
            TabIndex        =   137
            Top             =   1440
            Width           =   10635
            Begin VB.TextBox txtConvert 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4800
               MaxLength       =   50
               TabIndex        =   141
               Top             =   1260
               Width           =   855
            End
            Begin VB.TextBox txt_description_LateConvert 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4800
               MaxLength       =   50
               TabIndex        =   142
               Top             =   1620
               Width           =   3495
            End
            Begin VB.CommandButton Command4 
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
               TabIndex        =   139
               Top             =   120
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.TextBox txtFrom_LateConvert 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4800
               MaxLength       =   50
               TabIndex        =   138
               Top             =   540
               Width           =   855
            End
            Begin VB.TextBox txtTo_LateConvert 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4800
               MaxLength       =   50
               TabIndex        =   140
               Top             =   900
               Width           =   855
            End
            Begin VB.Label Label41 
               Caption         =   "Minutes"
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
               Left            =   5730
               TabIndex        =   150
               Top             =   1290
               Width           =   2205
            End
            Begin VB.Label Label40 
               Caption         =   "Minutes"
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
               Left            =   5730
               TabIndex        =   149
               Top             =   960
               Width           =   2205
            End
            Begin VB.Label Label35 
               Caption         =   "Minutes"
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
               Left            =   5730
               TabIndex        =   148
               Top             =   600
               Width           =   2205
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "CONVERTION*"
               Height          =   195
               Left            =   3495
               TabIndex        =   147
               Top             =   1290
               Width           =   1080
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "DESCRIPTION"
               Height          =   195
               Left            =   3570
               TabIndex        =   145
               Top             =   1620
               Width           =   1020
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "FROM*"
               Height          =   195
               Left            =   4050
               TabIndex        =   144
               Top             =   600
               Width           =   525
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "TO*"
               Height          =   195
               Left            =   4275
               TabIndex        =   143
               Top             =   930
               Width           =   300
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Data Control Button"
            Height          =   1335
            Left            =   -74820
            TabIndex        =   131
            Top             =   4020
            Width           =   10635
            Begin prj_tpc.vbButton cmdNew_LateConvert 
               Height          =   705
               Left            =   300
               TabIndex        =   132
               Top             =   360
               Width           =   945
               _extentx        =   1667
               _extenty        =   1244
               btype           =   14
               tx              =   "&New"
               enab            =   -1  'True
               font            =   "frm_mst_preferensi.frx":8658
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_mst_preferensi.frx":8684
               picn            =   "frm_mst_preferensi.frx":86A2
               umcol           =   -1  'True
               soft            =   0   'False
               picpos          =   2
               ngrey           =   0   'False
               fx              =   0
               hand            =   0   'False
               check           =   0   'False
               value           =   0   'False
            End
            Begin prj_tpc.vbButton cmdSave_LateConvert 
               Height          =   705
               Left            =   1320
               TabIndex        =   133
               Top             =   360
               Width           =   945
               _extentx        =   1667
               _extenty        =   1244
               btype           =   14
               tx              =   "&Save"
               enab            =   -1  'True
               font            =   "frm_mst_preferensi.frx":9736
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_mst_preferensi.frx":9762
               picn            =   "frm_mst_preferensi.frx":9780
               umcol           =   -1  'True
               soft            =   0   'False
               picpos          =   2
               ngrey           =   0   'False
               fx              =   0
               hand            =   0   'False
               check           =   0   'False
               value           =   0   'False
            End
            Begin prj_tpc.vbButton cmdEdit_LateConvert 
               Height          =   705
               Left            =   2340
               TabIndex        =   134
               Top             =   360
               Width           =   945
               _extentx        =   1667
               _extenty        =   1244
               btype           =   14
               tx              =   "&Edit"
               enab            =   -1  'True
               font            =   "frm_mst_preferensi.frx":A814
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_mst_preferensi.frx":A840
               picn            =   "frm_mst_preferensi.frx":A85E
               umcol           =   -1  'True
               soft            =   0   'False
               picpos          =   2
               ngrey           =   0   'False
               fx              =   0
               hand            =   0   'False
               check           =   0   'False
               value           =   0   'False
            End
            Begin prj_tpc.vbButton cmdDelete_LateConvert 
               Height          =   705
               Left            =   3360
               TabIndex        =   135
               Top             =   360
               Width           =   945
               _extentx        =   1667
               _extenty        =   1244
               btype           =   14
               tx              =   "&Delete"
               enab            =   -1  'True
               font            =   "frm_mst_preferensi.frx":B8F2
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_mst_preferensi.frx":B91E
               picn            =   "frm_mst_preferensi.frx":B93C
               umcol           =   -1  'True
               soft            =   0   'False
               picpos          =   2
               ngrey           =   0   'False
               fx              =   0
               hand            =   0   'False
               check           =   0   'False
               value           =   0   'False
            End
            Begin prj_tpc.vbButton cmdCancel_LateConvert 
               Height          =   705
               Left            =   4380
               TabIndex        =   136
               Top             =   360
               Width           =   945
               _extentx        =   1667
               _extenty        =   1244
               btype           =   14
               tx              =   "&Cancel"
               enab            =   -1  'True
               font            =   "frm_mst_preferensi.frx":C9D0
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_mst_preferensi.frx":C9FC
               picn            =   "frm_mst_preferensi.frx":CA1A
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
         Begin VB.Frame Frame10 
            Caption         =   "Data Control Button"
            Height          =   1335
            Left            =   -74820
            TabIndex        =   118
            Top             =   4050
            Width           =   10635
            Begin prj_tpc.vbButton cmdNew_PreAll 
               Height          =   705
               Left            =   300
               TabIndex        =   119
               Top             =   360
               Width           =   945
               _extentx        =   1667
               _extenty        =   1244
               btype           =   14
               tx              =   "&New"
               enab            =   -1  'True
               font            =   "frm_mst_preferensi.frx":DAAE
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_mst_preferensi.frx":DADA
               picn            =   "frm_mst_preferensi.frx":DAF8
               umcol           =   -1  'True
               soft            =   0   'False
               picpos          =   2
               ngrey           =   0   'False
               fx              =   0
               hand            =   0   'False
               check           =   0   'False
               value           =   0   'False
            End
            Begin prj_tpc.vbButton cmdSave_PreAll 
               Height          =   705
               Left            =   1320
               TabIndex        =   120
               Top             =   360
               Width           =   945
               _extentx        =   1667
               _extenty        =   1244
               btype           =   14
               tx              =   "&Save"
               enab            =   -1  'True
               font            =   "frm_mst_preferensi.frx":EB8C
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_mst_preferensi.frx":EBB8
               picn            =   "frm_mst_preferensi.frx":EBD6
               umcol           =   -1  'True
               soft            =   0   'False
               picpos          =   2
               ngrey           =   0   'False
               fx              =   0
               hand            =   0   'False
               check           =   0   'False
               value           =   0   'False
            End
            Begin prj_tpc.vbButton cmdEdit_PreAll 
               Height          =   705
               Left            =   2340
               TabIndex        =   121
               Top             =   360
               Width           =   945
               _extentx        =   1667
               _extenty        =   1244
               btype           =   14
               tx              =   "&Edit"
               enab            =   -1  'True
               font            =   "frm_mst_preferensi.frx":FC6A
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_mst_preferensi.frx":FC96
               picn            =   "frm_mst_preferensi.frx":FCB4
               umcol           =   -1  'True
               soft            =   0   'False
               picpos          =   2
               ngrey           =   0   'False
               fx              =   0
               hand            =   0   'False
               check           =   0   'False
               value           =   0   'False
            End
            Begin prj_tpc.vbButton cmdDelete_PreAll 
               Height          =   705
               Left            =   3360
               TabIndex        =   122
               Top             =   360
               Width           =   945
               _extentx        =   1667
               _extenty        =   1244
               btype           =   14
               tx              =   "&Delete"
               enab            =   -1  'True
               font            =   "frm_mst_preferensi.frx":10D48
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_mst_preferensi.frx":10D74
               picn            =   "frm_mst_preferensi.frx":10D92
               umcol           =   -1  'True
               soft            =   0   'False
               picpos          =   2
               ngrey           =   0   'False
               fx              =   0
               hand            =   0   'False
               check           =   0   'False
               value           =   0   'False
            End
            Begin prj_tpc.vbButton cmdCancel_PreAll 
               Height          =   705
               Left            =   4380
               TabIndex        =   123
               Top             =   360
               Width           =   945
               _extentx        =   1667
               _extenty        =   1244
               btype           =   14
               tx              =   "&Cancel"
               enab            =   -1  'True
               font            =   "frm_mst_preferensi.frx":11E26
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   15790320
               bcolo           =   15790320
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frm_mst_preferensi.frx":11E52
               picn            =   "frm_mst_preferensi.frx":11E70
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
         Begin VB.Frame fra_entry_PreAll 
            Height          =   2475
            Left            =   -74820
            TabIndex        =   115
            Top             =   1470
            Width           =   10635
            Begin VB.TextBox txtPercentage 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4800
               MaxLength       =   50
               TabIndex        =   126
               Top             =   1020
               Width           =   855
            End
            Begin VB.TextBox txtLimitHours 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4800
               MaxLength       =   50
               TabIndex        =   125
               Top             =   660
               Width           =   855
            End
            Begin VB.CommandButton Command3 
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
               TabIndex        =   116
               Top             =   120
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.TextBox txt_description_PreAll 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4800
               MaxLength       =   50
               TabIndex        =   128
               Top             =   1380
               Width           =   3495
            End
            Begin VB.Label Label33 
               Caption         =   "% OF PRESENCE ALLOWANCE"
               Height          =   225
               Left            =   5730
               TabIndex        =   130
               Top             =   1080
               Width           =   2505
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "EXPENSE*"
               Height          =   195
               Left            =   3840
               TabIndex        =   129
               Top             =   1050
               Width           =   735
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "LIMIT HOURS*"
               Height          =   195
               Left            =   3510
               TabIndex        =   127
               Top             =   720
               Width           =   1065
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "DESCRIPTION"
               Height          =   195
               Left            =   3570
               TabIndex        =   117
               Top             =   1380
               Width           =   1020
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_PreAll 
            Height          =   3465
            Left            =   -74820
            TabIndex        =   124
            Top             =   480
            Width           =   10635
            _ExtentX        =   18759
            _ExtentY        =   6112
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
            Columns(1).Caption=   "LIMIT HOURS"
            Columns(1).DataField=   "limit_hours"
            Columns(1).NumberFormat=   "Standard"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "EXPENSE (%)"
            Columns(2).DataField=   "percentage"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=953"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=873"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3757"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3678"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=3122"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3043"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=9922"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=9843"
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
            Caption         =   "LIST PRESENCE ALLOWANCE"
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
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
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
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_LateConvert 
            Height          =   3465
            Left            =   -74820
            TabIndex        =   146
            Top             =   450
            Width           =   10635
            _ExtentX        =   18759
            _ExtentY        =   6112
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
            Columns(1).Caption=   "FROM (MINUTES)"
            Columns(1).DataField=   "from_value"
            Columns(1).NumberFormat=   "Standard"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "TO (MINUTES)"
            Columns(2).DataField=   "to_value"
            Columns(2).NumberFormat=   "Standard"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "CONVERTION (MINUTES)"
            Columns(3).DataField=   "convert_value"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=953"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=873"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3069"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2990"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=3016"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2937"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2672"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2593"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=8017"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=7938"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
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
            Caption         =   "LIST TIME LATE CONVERTION"
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
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
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
         Begin prj_tpc.vbButton cmdSave_Gen 
            Height          =   705
            Left            =   9480
            TabIndex        =   173
            Top             =   4590
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":12F04
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":12F30
            picn            =   "frm_mst_preferensi.frx":12F4E
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
         Height          =   3255
         Left            =   -73140
         TabIndex        =   99
         Top             =   840
         Width           =   7965
         Begin VB.CheckBox chkPLConvert 
            Caption         =   "NO"
            Height          =   285
            Left            =   3210
            TabIndex        =   224
            Top             =   2820
            Width           =   1455
         End
         Begin VB.TextBox txt_jstk_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   4830
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   179
            Top             =   2490
            Width           =   2535
         End
         Begin VB.TextBox txt_ptkp_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   4830
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   178
            Top             =   2070
            Width           =   2535
         End
         Begin VB.TextBox txt_pph21_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   4830
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   177
            Top             =   1650
            Width           =   2535
         End
         Begin VB.TextBox txt_company_name_gen 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   4785
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   175
            Top             =   330
            Width           =   2565
         End
         Begin VB.TextBox txt_value_wh 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3210
            TabIndex        =   103
            Top             =   1170
            Width           =   1545
         End
         Begin VB.CheckBox chkPunishment 
            Caption         =   "YES"
            Height          =   225
            Left            =   6600
            TabIndex        =   102
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtLateTolerance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3210
            TabIndex        =   101
            Top             =   810
            Width           =   735
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_company_gen 
            Height          =   375
            Left            =   3210
            OleObjectBlob   =   "frm_mst_preferensi.frx":13FE2
            TabIndex        =   100
            Top             =   330
            Width           =   1545
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_jstk 
            Height          =   375
            Left            =   3210
            OleObjectBlob   =   "frm_mst_preferensi.frx":15F4C
            TabIndex        =   106
            Top             =   2490
            Width           =   1575
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_ptkp 
            Height          =   375
            Left            =   3210
            OleObjectBlob   =   "frm_mst_preferensi.frx":17EB7
            TabIndex        =   105
            Top             =   2070
            Width           =   1575
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pph 
            Height          =   375
            Left            =   3210
            OleObjectBlob   =   "frm_mst_preferensi.frx":19E12
            TabIndex        =   104
            Top             =   1650
            Width           =   1575
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NOT PRSENCE CONVERTED TO PL*"
            Height          =   195
            Left            =   480
            TabIndex        =   223
            Top             =   2850
            Width           =   2535
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JAMSOSTEK TYPE*"
            Height          =   195
            Left            =   1530
            TabIndex        =   182
            Top             =   2520
            Width           =   1485
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PTKP TYPE*"
            Height          =   195
            Left            =   2070
            TabIndex        =   181
            Top             =   2100
            Width           =   945
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PPh21 TYPE*"
            Height          =   195
            Left            =   2010
            TabIndex        =   180
            Top             =   1680
            Width           =   1005
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DEFAULT COMPANY*"
            Height          =   195
            Left            =   1485
            TabIndex        =   176
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TOTAL WORKING HOURS / MONTH*"
            Height          =   195
            Left            =   420
            TabIndex        =   174
            Top             =   1200
            Width           =   2610
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "PUNISHMENT"
            Height          =   285
            Left            =   4770
            TabIndex        =   109
            Top             =   840
            Width           =   1665
         End
         Begin VB.Label Label29 
            Caption         =   "Minutes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4080
            TabIndex        =   108
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "LATE TOLERANCE"
            Height          =   285
            Left            =   1380
            TabIndex        =   107
            Top             =   840
            Width           =   1665
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   1890
         TabIndex        =   69
         Top             =   4800
         Width           =   7965
         Begin prj_tpc.vbButton cmdSave_OT 
            Height          =   705
            Left            =   5670
            TabIndex        =   70
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":1BD6C
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":1BD98
            picn            =   "frm_mst_preferensi.frx":1BDB6
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
         Left            =   -73140
         TabIndex        =   62
         Top             =   4200
         Width           =   7965
         Begin prj_tpc.vbButton cmdSave_Att 
            Height          =   705
            Left            =   5820
            TabIndex        =   63
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":1CE4A
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":1CE76
            picn            =   "frm_mst_preferensi.frx":1CE94
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
      Begin VB.Frame fra_entry_jhk 
         Height          =   2865
         Left            =   -74850
         TabIndex        =   44
         Top             =   1980
         Width           =   11295
         Begin VB.CheckBox chkAllDept_Presence 
            Caption         =   "ALL DEPARTMENT"
            Height          =   225
            Left            =   3060
            TabIndex        =   50
            Top             =   1170
            Width           =   1725
         End
         Begin VB.TextBox txt_description_presence 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   49
            Top             =   2160
            Width           =   3495
         End
         Begin VB.TextBox txt_value_jhk 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            TabIndex        =   48
            Top             =   1800
            Width           =   1545
         End
         Begin VB.TextBox txt_department_name_presence 
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
            TabIndex        =   47
            Top             =   1440
            Width           =   3855
         End
         Begin VB.TextBox txt_company_name_presence 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   4620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   46
            Top             =   750
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
            TabIndex        =   45
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_company_presence 
            Height          =   375
            Left            =   3060
            OleObjectBlob   =   "frm_mst_preferensi.frx":1DF28
            TabIndex        =   51
            Top             =   750
            Width           =   1545
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_department_presence 
            Height          =   375
            Left            =   3060
            OleObjectBlob   =   "frm_mst_preferensi.frx":1FE97
            TabIndex        =   52
            Top             =   1440
            Width           =   1545
         End
         Begin MSComCtl2.DTPicker DTPicker_Presence 
            Height          =   315
            Left            =   3060
            TabIndex        =   58
            Top             =   390
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96796675
            CurrentDate     =   41213
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EFFECTIVE DATE*"
            Height          =   195
            Left            =   1470
            TabIndex        =   59
            Top             =   420
            Width           =   1410
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "COMPANY*"
            Height          =   195
            Left            =   2025
            TabIndex        =   56
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DEPARTMENT"
            Height          =   195
            Left            =   1890
            TabIndex        =   55
            Top             =   1170
            Width           =   990
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TOTAL WORKING DAYS / MONTH*"
            Height          =   195
            Left            =   375
            TabIndex        =   54
            Top             =   1830
            Width           =   2475
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   1830
            TabIndex        =   53
            Top             =   2190
            Width           =   1020
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74850
         TabIndex        =   38
         Top             =   4830
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew_Presence 
            Height          =   705
            Left            =   300
            TabIndex        =   39
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&New"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":21E09
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":21E35
            picn            =   "frm_mst_preferensi.frx":21E53
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_Presence 
            Height          =   705
            Left            =   1320
            TabIndex        =   40
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":22EE7
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":22F13
            picn            =   "frm_mst_preferensi.frx":22F31
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_Presence 
            Height          =   705
            Left            =   2340
            TabIndex        =   41
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Edit"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":23FC5
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":23FF1
            picn            =   "frm_mst_preferensi.frx":2400F
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_Presence 
            Height          =   705
            Left            =   3360
            TabIndex        =   42
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Delete"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":250A3
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":250CF
            picn            =   "frm_mst_preferensi.frx":250ED
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_Presence 
            Height          =   705
            Left            =   4380
            TabIndex        =   43
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Cancel"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":26181
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":261AD
            picn            =   "frm_mst_preferensi.frx":261CB
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
      Begin VB.Frame Frame3 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74910
         TabIndex        =   31
         Top             =   4830
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew_Leave 
            Height          =   705
            Left            =   240
            TabIndex        =   32
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&New"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":2725F
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":2728B
            picn            =   "frm_mst_preferensi.frx":272A9
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_Leave 
            Height          =   705
            Left            =   1260
            TabIndex        =   33
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":2833D
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":28369
            picn            =   "frm_mst_preferensi.frx":28387
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_Leave 
            Height          =   705
            Left            =   2280
            TabIndex        =   34
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Edit"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":2941B
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":29447
            picn            =   "frm_mst_preferensi.frx":29465
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_Leave 
            Height          =   705
            Left            =   3300
            TabIndex        =   35
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Delete"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":2A4F9
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":2A525
            picn            =   "frm_mst_preferensi.frx":2A543
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_Leave 
            Height          =   705
            Left            =   4320
            TabIndex        =   36
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Cancel"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":2B5D7
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":2B603
            picn            =   "frm_mst_preferensi.frx":2B621
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
      Begin VB.Frame fra_entry_leave 
         Height          =   2895
         Left            =   -74910
         TabIndex        =   18
         Top             =   1830
         Width           =   11295
         Begin VB.TextBox txt_department_name_leave 
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
            TabIndex        =   24
            Top             =   1380
            Width           =   3855
         End
         Begin VB.TextBox txt_company_name_leave 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   4920
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   23
            Top             =   690
            Width           =   3855
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
            TabIndex        =   22
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.ComboBox cboLeave 
            Height          =   315
            ItemData        =   "frm_mst_preferensi.frx":2C6B5
            Left            =   3360
            List            =   "frm_mst_preferensi.frx":2C6BF
            TabIndex        =   21
            Text            =   "PERIODE KERJA"
            Top             =   1740
            Width           =   1545
         End
         Begin VB.TextBox txt_description_leave 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3360
            MaxLength       =   50
            TabIndex        =   20
            Top             =   2100
            Width           =   5415
         End
         Begin VB.CheckBox chkAllDept_Leave 
            Caption         =   "ALL DEPARTMENT"
            Height          =   225
            Left            =   3360
            TabIndex        =   19
            Top             =   1110
            Width           =   1725
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_company_leave 
            Height          =   375
            Left            =   3360
            OleObjectBlob   =   "frm_mst_preferensi.frx":2C6D9
            TabIndex        =   25
            Top             =   690
            Width           =   1545
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_department_leave 
            Height          =   375
            Left            =   3360
            OleObjectBlob   =   "frm_mst_preferensi.frx":2E645
            TabIndex        =   26
            Top             =   1380
            Width           =   1545
         End
         Begin MSComCtl2.DTPicker DTPicker_Leave 
            Height          =   315
            Left            =   3360
            TabIndex        =   60
            Top             =   330
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96796675
            CurrentDate     =   41213
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EFFECTIVE DATE*"
            Height          =   195
            Left            =   1770
            TabIndex        =   61
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PERUSAHAAN"
            Height          =   195
            Left            =   2145
            TabIndex        =   30
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DEPARTMENT"
            Height          =   195
            Left            =   2160
            TabIndex        =   29
            Top             =   1110
            Width           =   990
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TYPE PERIODE CUTI"
            Height          =   195
            Left            =   1695
            TabIndex        =   28
            Top             =   1770
            Width           =   1470
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   1920
            TabIndex        =   27
            Top             =   2100
            Width           =   1230
         End
      End
      Begin VB.Frame fra_entry 
         Height          =   2655
         Left            =   -74850
         TabIndex        =   9
         Top             =   2040
         Width           =   11295
         Begin VB.TextBox txt_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   12
            Top             =   1500
            Width           =   3495
         End
         Begin VB.TextBox txt_umk_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1140
            Width           =   3495
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
            TabIndex        =   10
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin MSComCtl2.DTPicker DTPicker_umk 
            Height          =   315
            Left            =   4800
            TabIndex        =   13
            Top             =   750
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   96796675
            CurrentDate     =   41213
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   3570
            TabIndex        =   16
            Top             =   1500
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EFFECTIVE DATE*"
            Height          =   195
            Left            =   3270
            TabIndex        =   15
            Top             =   780
            Width           =   1320
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UMK VALUE*"
            Height          =   195
            Left            =   3660
            TabIndex        =   14
            Top             =   1140
            Width           =   915
         End
      End
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74850
         TabIndex        =   3
         Top             =   4800
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew_UMK 
            Height          =   705
            Left            =   300
            TabIndex        =   4
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&New"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":305B4
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":305E0
            picn            =   "frm_mst_preferensi.frx":305FE
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSave_UMK 
            Height          =   705
            Left            =   1320
            TabIndex        =   5
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Save"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":31692
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":316BE
            picn            =   "frm_mst_preferensi.frx":316DC
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEdit_UMK 
            Height          =   705
            Left            =   2340
            TabIndex        =   6
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Edit"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":32770
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":3279C
            picn            =   "frm_mst_preferensi.frx":327BA
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDelete_UMK 
            Height          =   705
            Left            =   3360
            TabIndex        =   7
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Delete"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":3384E
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":3387A
            picn            =   "frm_mst_preferensi.frx":33898
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCancel_UMK 
            Height          =   705
            Left            =   4380
            TabIndex        =   8
            Top             =   360
            Width           =   945
            _extentx        =   1667
            _extenty        =   1244
            btype           =   14
            tx              =   "&Cancel"
            enab            =   -1  'True
            font            =   "frm_mst_preferensi.frx":3492C
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   15790320
            bcolo           =   15790320
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frm_mst_preferensi.frx":34958
            picn            =   "frm_mst_preferensi.frx":34976
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_UMK 
         Height          =   4245
         Left            =   -74850
         TabIndex        =   17
         Top             =   450
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "ADD DATE"
         Columns(0).DataField=   "entry_date"
         Columns(0).NumberFormat=   "yyyy-MM-dd"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "EFFECTIVE DATE"
         Columns(1).DataField=   "effective_date"
         Columns(1).NumberFormat=   "yyyy-MM-dd"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "UMK"
         Columns(2).DataField=   "umk_value"
         Columns(2).NumberFormat=   "Standard"
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
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=4180"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4101"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=9260"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=9181"
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
         Caption         =   "LIST UMK"
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
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Leave 
         Height          =   4245
         Left            =   -74910
         TabIndex        =   37
         Top             =   480
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "ADD DATE"
         Columns(0).DataField=   "entry_date"
         Columns(0).NumberFormat=   "yyyy-MM-dd"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "EFFECTIVE DATE"
         Columns(1).DataField=   "effective_date"
         Columns(1).NumberFormat=   "yyyy-MM-dd"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "COMPANY CODE"
         Columns(2).DataField=   "company_code"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "COMPANY"
         Columns(3).DataField=   "company_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "FLAG DEPT."
         Columns(4).DataField=   "flag_dept"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DEPT. CODE"
         Columns(5).DataField=   "department_code"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DEPARTMENT"
         Columns(6).DataField=   "department_name"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "LEAVE TYPE"
         Columns(7).DataField=   "leave_type"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "TYPE"
         Columns(8).DataField=   "type"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "DESCRIPTION"
         Columns(9).DataField=   "description"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2064"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2302"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2223"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=3387"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=3307"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=2963"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2884"
         Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(32)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=3466"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=3387"
         Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(39)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(43)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(45)=   "Column(8).Width=2408"
         Splits(0)._ColumnProps(46)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(8)._WidthInPix=2328"
         Splits(0)._ColumnProps(48)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(49)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(50)=   "Column(9).Width=5265"
         Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=5186"
         Splits(0)._ColumnProps(53)=   "Column(9)._ColStyle=516"
         Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
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
         Caption         =   "LIST OF LEAVE"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=70,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=28,.parent=13"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Presence 
         Height          =   4245
         Left            =   -74850
         TabIndex        =   57
         Top             =   600
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "ADD DATE"
         Columns(0).DataField=   "entry_date"
         Columns(0).NumberFormat=   "yyyy-MM-dd"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "EFFECTIVE DATE"
         Columns(1).DataField=   "effective_date"
         Columns(1).NumberFormat=   "yyyy-MM-dd"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "COMPANY CODE"
         Columns(2).DataField=   "company_code"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "COMPANY"
         Columns(3).DataField=   "company_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "FLAG DEPT"
         Columns(4).DataField=   "flag_dept"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DEPT. CODE"
         Columns(5).DataField=   "department_code"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DEPARTMENT"
         Columns(6).DataField=   "department_name"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "TOT. WORK DAYS"
         Columns(7).DataField=   "jhk_value"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "DESCRIPTION"
         Columns(8).DataField=   "description"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2540"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2461"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=4789"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=4710"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=2963"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2884"
         Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(32)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=3201"
         Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(39)=   "Column(7).Width=2434"
         Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2355"
         Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=513"
         Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(44)=   "Column(8).Width=3784"
         Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=3704"
         Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
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
         Caption         =   "LIST OF PRESENCE"
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
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=70,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=28,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=17"
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
   End
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   10560
      TabIndex        =   2
      Top             =   7230
      Width           =   945
      _extentx        =   1667
      _extenty        =   1244
      btype           =   14
      tx              =   "&Exit"
      enab            =   -1  'True
      font            =   "frm_mst_preferensi.frx":35A0A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   15790320
      bcolo           =   15790320
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frm_mst_preferensi.frx":35A36
      picn            =   "frm_mst_preferensi.frx":35A54
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER PREFERENCE"
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
      TabIndex        =   0
      Top             =   150
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_preferensi.frx":36AE8
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
Dim rsAtt As New ADODB.Recordset
Dim rsot As New ADODB.Recordset
Dim rsOT_TDB As New ADODB.Recordset

Dim rsPreAll As New ADODB.Recordset
Dim rsLateConvert As New ADODB.Recordset
Dim rsPenaltyGeneral As New ADODB.Recordset
Dim rsGen As New ADODB.Recordset
Dim rsPeriode As New ADODB.Recordset

Dim rsPPh As New ADODB.Recordset
Dim rsPTKP As New ADODB.Recordset
Dim rsJSTK As New ADODB.Recordset

Dim vIdNumber As Integer
            
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Public public_int_mode As Integer

Private Function check_validate_exist_new() As Boolean
Dim str_sql As String
    check_validate_exist_new = False
    
    If SSTab1.Tab = 2 Then
        If rs.State Then rs.Close
        str_sql = "select count(company_code) as rec_count from m_pref_presence " & _
                    "where company_code = '" & TDBCombo_company_presence.Text & "' and department_code = '" & TDBCombo_department_presence.Text & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
        
        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            rs.Close
            Exit Function
        End If
        
        rs.Close
    ElseIf SSTab1.Tab = 3 Then
        If rs.State Then rs.Close
        str_sql = "select count(effective_date) as rec_count from m_pref_umk " & _
                    "where date(effective_date) = '" & Format(DTPicker_umk.Value, "yyyy-MM-dd") & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
        
        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            rs.Close
            Exit Function
        End If
        
        rs.Close
    ElseIf SSTab1.Tab = 4 Then
        If rs.State Then rs.Close
        str_sql = "select count(company_code) as rec_count from m_pref_leave " & _
                    "where company_code = '" & TDBCombo_company_leave.Text & "' and department_code = '" & TDBCombo_department_leave.Text & "'"
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
    If SSTab1.Tab = 2 Then
        MsgBox "Data Sudah Ada...", vbCritical, headerMSG
        DTPicker_Presence.Value = Now
        TDBCombo_company_presence.Text = ""
        If TDBCombo_company_presence.Enabled = True Then TDBCombo_company_presence.SetFocus
    ElseIf SSTab1.Tab = 3 Then
        MsgBox "Data Sudah Ada...", vbCritical, headerMSG
        DTPicker_umk.Value = Now
        If DTPicker_umk.Enabled = True Then txt_umk_value.SetFocus
    ElseIf SSTab1.Tab = 4 Then
        MsgBox "Data Sudah Ada...", vbCritical, headerMSG
        DTPicker_Leave.Value = Now
        TDBCombo_company_leave.Text = ""
        If TDBCombo_company_leave.Enabled = True Then TDBCombo_company_leave.SetFocus
    End If
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If SSTab1.Tab = 2 Then
        If Not TDBCombo_company_presence.Text = rsJHK.Fields("company_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 3 Then
        If Not DTPicker_umk.Value = rsUMK.Fields("effective_date").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 4 Then
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
        If Trim(TDBCombo_company_gen.Text) = "" Then
            MsgBox "Company Is Empty...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_company_presence.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If Trim(TDBCombo_pph.Text) = "" Then
            MsgBox "PPh21 Is Empty...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_company_presence.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If Trim(TDBCombo_ptkp.Text) = "" Then
            MsgBox "PTKP Is Empty...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_company_presence.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If Trim(TDBCombo_jstk.Text) = "" Then
            MsgBox "Jamsostek Is Empty...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_company_presence.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        'validasi tdb_company
        If Trim(TDBCombo_company_presence.Text) = "" Then
            MsgBox "Perusahaan Is Empty...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_company_presence.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If chkAllDept_Presence.Value = 0 Then
            'validasi tdb_division
            If Trim(TDBCombo_department_presence.Text) = "" Then
                MsgBox "Divisi Is Empty...", vbOKOnly + vbInformation, headerMSG
                TDBCombo_department_presence.SetFocus
                check_validate_new = False
                Exit Function
            End If
        End If
        
        'validasi jhk_value
        If Trim(txt_value_jhk.Text) = "" Then
            MsgBox "JHK Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_value_jhk.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 3 Then
        'validasi umk
        If Trim(txt_umk_value) = "" Then
            MsgBox "Nilai UMK Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_umk_value.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 4 Then
        'validasi tdb_company
        If Trim(TDBCombo_company_leave.Text) = "" Then
            MsgBox "Perusahaan Is Empty...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_company_leave.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        If chkAllDept_Leave.Value = 0 Then
            'validasi tdb_division
            If Trim(TDBCombo_department_leave.Text) = "" Then
                MsgBox "Divisi Is Empty...", vbOKOnly + vbInformation, headerMSG
                TDBCombo_department_leave.SetFocus
                check_validate_new = False
                Exit Function
            End If
        End If
    ElseIf SSTab1.Tab = 5 Then
        If SSTab2.Tab = 0 Then
            'validasi txtLimitHours
            If Trim(txtLimitHours.Text) = "" Then
                MsgBox "Limit Hours Is Empty...", vbOKOnly + vbInformation, headerMSG
                txtLimitHours.SetFocus
                check_validate_new = False
                Exit Function
            End If
            
            'validasi txtPercentage
            If Trim(txtPercentage.Text) = "" Then
                MsgBox "Percentage Is Empty...", vbOKOnly + vbInformation, headerMSG
                txtPercentage.SetFocus
                check_validate_new = False
                Exit Function
            End If
        ElseIf SSTab2.Tab = 1 Then
            'validasi txtFrom
            If Trim(txtFrom_LateConvert.Text) = "" Then
                MsgBox "From Value Is Empty...", vbOKOnly + vbInformation, headerMSG
                txtFrom_LateConvert.SetFocus
                check_validate_new = False
                Exit Function
            End If
            
            'validasi txtTo
            If Trim(txtTo_LateConvert.Text) = "" Then
                MsgBox "To Value Is Empty...", vbOKOnly + vbInformation, headerMSG
                txtTo_LateConvert.SetFocus
                check_validate_new = False
                Exit Function
            End If
            
            'validasi txtConvert
            If Trim(txtConvert.Text) = "" Then
                MsgBox "Convert Value Is Empty...", vbOKOnly + vbInformation, headerMSG
                txtConvert.SetFocus
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
    Dim vEffectiveDate As String
    
    If SSTab1.Tab = 2 Then
        If Not (TDBGrid_Presence.ApproxCount > 0 And TDBGrid_Presence.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Are You Sure to Delete '" _
            & TDBGrid_Presence.Columns("company_name").Value & "' - '" & TDBGrid_Presence.Columns("department_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        vCompany = TDBGrid_Presence.Columns("company_code").Value
        vDivision = TDBGrid_Presence.Columns("department_code").Value
        vEffectiveDate = Format(TDBGrid_Presence.Columns("effective_date").Value, "yyyy-MM-dd")
        
        CnG.Execute "delete from m_pref_presence " & _
            "where company_code = '" & vCompany & "' and department_code = '" & vDivision & "' and date(effective_date) = '" & vEffectiveDate & "'"
        CnG.CommitTrans
    ElseIf SSTab1.Tab = 3 Then
        If Not (TDBGrid_UMK.ApproxCount > 0 And TDBGrid_UMK.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Are You Sure to Delete '" _
            & TDBGrid_UMK.Columns("effective_date").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        
        vEffectiveDate = Format(TDBGrid_UMK.Columns("effective_date").Value, "yyyy-MM-dd")
        
        CnG.Execute "delete from m_pref_umk where date(effective_date) = '" _
            & Format(vEffectiveDate, "yyyy-MM-dd") & "'"
        CnG.CommitTrans
    ElseIf SSTab1.Tab = 4 Then
        If Not (TDBGrid_Leave.ApproxCount > 0 And TDBGrid_Leave.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Are You Sure to Delete '" _
            & TDBGrid_Leave.Columns("company_name").Value & "' - '" & TDBGrid_Leave.Columns("department_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        vCompany = TDBGrid_Leave.Columns("company_code").Value
        vDivision = TDBGrid_Leave.Columns("department_code").Value
        vEffectiveDate = Format(TDBGrid_Leave.Columns("effective_date").Value, "yyyy-MM-dd")
        
        CnG.Execute "delete from m_pref_leave " & _
            "where company_code = '" & vCompany & "' and department_code = '" & vDivision & "' and date(effective_date) = '" & vEffectiveDate & "'"
        CnG.CommitTrans
    ElseIf SSTab1.Tab = 5 Then
        If SSTab2.Tab = 0 Then
            If Not (TDBGrid_PreAll.ApproxCount > 0 And TDBGrid_PreAll.Bookmark > 0) Then
                MsgBox "There Is No Data Selected...", vbInformation, headerMSG
                Exit Sub
            End If
            
            i = MsgBox("Are You Sure to Delete '" _
                & TDBGrid_PreAll.Columns("limit_hours").Value & "' - '" & TDBGrid_Leave.Columns("department_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
            If Not i = vbYes Then Exit Sub
            
            CnG.BeginTrans
            
            CnG.Execute "delete from m_pref_preall " & _
                "where id_number = '" & TDBGrid_PreAll.Columns("id_number").Value & "'"
            CnG.CommitTrans
        ElseIf SSTab2.Tab = 1 Then
            If Not (TDBGrid_LateConvert.ApproxCount > 0 And TDBGrid_LateConvert.Bookmark > 0) Then
                MsgBox "There Is No Data Selected...", vbInformation, headerMSG
                Exit Sub
            End If
            
            i = MsgBox("Are You Sure to Delete '" _
                & TDBGrid_LateConvert.Columns("convert_value").Value & "' - '" & TDBGrid_Leave.Columns("department_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
            If Not i = vbYes Then Exit Sub
            
            CnG.BeginTrans
            
            CnG.Execute "delete from m_pref_lateconvert " & _
                "where id_number = '" & TDBGrid_LateConvert.Columns("id_number").Value & "'"
            CnG.CommitTrans
        End If
    End If
    
    Call load_data
    int_mode = 0
    Call load_mode
End Sub

Public Sub set_edit_data()
On Error GoTo Err

Dim vFlagOTMethod As Integer
Dim vFlagCalcStart As Integer
Dim v_flag_presence As Integer
Dim v_flag_meal As Integer
Dim v_flag_transport As Integer

Dim vFlagPeriode As Integer

    vSetData = 1
    
    If SSTab1.Tab = 0 Then
        If rsAtt.RecordCount > 0 Then
            With rsAtt
                TDBCombo_company_gen.Text = IIf(IsNull(.Fields("company_code").Value), "", .Fields("company_code").Value)
                txt_company_name_gen.Text = IIf(IsNull(.Fields("company_name").Value), "", .Fields("company_name").Value)
                
                txtLateTolerance.Text = IIf(IsNull(.Fields("late_tolerance").Value), 0, .Fields("late_tolerance").Value)
                chkPunishment.Value = IIf(IsNull(.Fields("flag_punishment").Value), 0, .Fields("flag_punishment").Value)
                txt_value_wh.Text = IIf(IsNull(.Fields("wh_value").Value), 0, .Fields("wh_value").Value)
                
                TDBCombo_pph.Text = IIf(IsNull(.Fields("pph21_code").Value), "", .Fields("pph21_code").Value)
                txt_pph21_name.Text = IIf(IsNull(.Fields("pph21_name").Value), "", .Fields("pph21_name").Value)
                TDBCombo_ptkp.Text = IIf(IsNull(.Fields("ptkp_code").Value), "", .Fields("ptkp_code").Value)
                txt_ptkp_name.Text = IIf(IsNull(.Fields("ptkp_name").Value), "", .Fields("ptkp_name").Value)
                TDBCombo_jstk.Text = IIf(IsNull(.Fields("jamsostek_code").Value), "", .Fields("jamsostek_code").Value)
                txt_jstk_name.Text = IIf(IsNull(.Fields("jamsostek_name").Value), "", .Fields("jamsostek_name").Value)
                
                chkPLConvert.Value = IIf(IsNull(.Fields("flag_pl").Value), 1, .Fields("flag_pl").Value)
            End With
        End If
    ElseIf SSTab1.Tab = 1 Then
        vSetData = 0
        If rsot.RecordCount > 0 Then
            With rsot
                vFlagOTMethod = IIf(IsNull(.Fields("flag_ot_method").Value), 0, .Fields("flag_ot_method").Value)
                chk_approve.Value = IIf(IsNull(.Fields("flag_auto_approve").Value), 0, .Fields("flag_auto_approve").Value)
                optSPL.Value = IIf(vFlagOTMethod = 0, 1, 0)
                optSystem.Value = IIf(vFlagOTMethod = 0, 0, 1)
                txtMAXOt(0).Text = IIf(IsNull(.Fields("ot_max_meal").Value), 0, .Fields("ot_max_meal").Value)
                chk_meal.Value = IIf(IsNull(.Fields("flag_meal").Value), 0, .Fields("flag_meal").Value)
                txtMAXOt(1).Text = IIf(IsNull(.Fields("ot_max_trans").Value), 0, .Fields("ot_max_trans").Value)
                chk_transport.Value = IIf(IsNull(.Fields("flag_trans").Value), 0, .Fields("flag_trans").Value)
                txtAutoOt_Sat.Text = IIf(IsNull(.Fields("auto_ot_sat").Value), 0, .Fields("auto_ot_sat").Value)
                txtAutoOt_wd.Text = IIf(IsNull(.Fields("auto_ot_wd").Value), 0, .Fields("auto_ot_wd").Value)
                
                vFlagCalcStart = IIf(IsNull(.Fields("flag_calc_start").Value), 0, .Fields("flag_calc_start").Value)
                optAfter.Value = IIf(vFlagCalcStart = 0, 1, 0)
                optDelay.Value = IIf(vFlagCalcStart = 0, 0, 1)
                txtOTStart.Text = IIf(IsNull(.Fields("delay_hours").Value), 0, .Fields("delay_hours").Value)
                TDBCombo_ot1.Text = IIf(IsNull(.Fields("ot_type").Value), "", .Fields("ot_type").Value)
                txt_ot_name1.Text = IIf(IsNull(.Fields("ot_name1").Value), "", .Fields("ot_name1").Value)
                TDBCombo_ot2.Text = IIf(IsNull(.Fields("ot_sat").Value), "", .Fields("ot_sat").Value)
                txt_ot_name2.Text = IIf(IsNull(.Fields("ot_name2").Value), "", .Fields("ot_name2").Value)
                TDBCombo_ot3.Text = IIf(IsNull(.Fields("ot_wd").Value), "", .Fields("ot_wd").Value)
                txt_ot_name3.Text = IIf(IsNull(.Fields("ot_name3").Value), "", .Fields("ot_name3").Value)
                
                txt_on_call_allowance.Text = FormatNumber(IIf(IsNull(.Fields("on_call_allowance").Value), 0, .Fields("on_call_allowance").Value))
                txt_change_status.Text = IIf(IsNull(.Fields("change_status_hours").Value), 0, .Fields("change_status_hours").Value)
            End With
        End If
        
'        txtMAXOt(0).SetFocus
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_Presence.ApproxCount > 0 And TDBGrid_Presence.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsJHK
            DTPicker_Presence.Value = .Fields("effective_date").Value
            TDBCombo_company_presence.Text = .Fields("company_code").Value
            txt_company_name_presence.Text = .Fields("company_name").Value
            chkAllDept_Presence.Value = IIf(IsNull(.Fields("flag_dept").Value), 0, .Fields("flag_dept").Value)
            TDBCombo_department_presence.Text = IIf(IsNull(.Fields("department_code").Value), "", .Fields("department_code").Value)
            txt_department_name_presence.Text = IIf(IsNull(.Fields("department_name").Value), "", .Fields("department_name").Value)
            txt_value_jhk.Text = .Fields("jhk_value").Value
            txt_description_presence.Text = .Fields("description").Value
        End With
    ElseIf SSTab1.Tab = 3 Then
        If Not (TDBGrid_UMK.ApproxCount > 0 And TDBGrid_UMK.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsUMK
            DTPicker_umk.Value = .Fields("effective_date").Value
            txt_umk_value = FormatNumber(.Fields("umk_value").Value)
            txt_description = .Fields("description").Value
        End With
    ElseIf SSTab1.Tab = 4 Then
        If Not (TDBGrid_Leave.ApproxCount > 0 And TDBGrid_Leave.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsLeave
            DTPicker_Leave.Value = .Fields("effective_date").Value
            TDBCombo_company_leave.Text = .Fields("company_code").Value
            txt_company_name_leave.Text = .Fields("company_name").Value
            chkAllDept_Leave.Value = IIf(IsNull(.Fields("flag_dept").Value), 0, .Fields("flag_dept").Value)
            TDBCombo_department_leave.Text = IIf(IsNull(.Fields("department_code").Value), "", .Fields("department_code").Value)
            txt_department_name_leave.Text = IIf(IsNull(.Fields("department_name").Value), "", .Fields("department_name").Value)
            cboLeave.ListIndex = .Fields("leave_type").Value
            txt_description_leave.Text = .Fields("description").Value
        End With
    ElseIf SSTab1.Tab = 5 Then
        If SSTab2.Tab = 0 Then
            If Not (TDBGrid_PreAll.ApproxCount > 0 And TDBGrid_PreAll.Bookmark > 0) Then
                MsgBox "There Is No Data Selected...", vbInformation, headerMSG
                vSetData = 0
                Exit Sub
            End If
            
            With rsPreAll
               vIdNumber = .Fields("id_number").Value
               txtLimitHours.Text = .Fields("limit_hours").Value
               txtPercentage.Text = .Fields("percentage").Value
               txt_description_PreAll.Text = .Fields("description").Value
            End With
        ElseIf SSTab2.Tab = 1 Then
            If Not (TDBGrid_LateConvert.ApproxCount > 0 And TDBGrid_LateConvert.Bookmark > 0) Then
                MsgBox "There Is No Data Selected...", vbInformation, headerMSG
                vSetData = 0
                Exit Sub
            End If
            
            With rsLateConvert
               vIdNumber = .Fields("id_number").Value
               txtFrom_LateConvert.Text = .Fields("from_value").Value
               txtTo_LateConvert.Text = .Fields("to_value").Value
               txtConvert.Text = .Fields("convert_value").Value
               txt_description_LateConvert.Text = .Fields("description").Value
            End With
        ElseIf SSTab2.Tab = 2 Then
            If rsPenaltyGeneral.RecordCount > 0 Then
                With rsPenaltyGeneral
                    txtPL.Text = IIf(IsNull(.Fields("pengali_pl").Value), 0, .Fields("pengali_pl").Value)
                    chkPL.Value = IIf(IsNull(.Fields("flag_pot_pl").Value), 0, .Fields("flag_pot_pl").Value)
                    txtPL_Not.Text = IIf(IsNull(.Fields("pengali_pl_not").Value), 0, .Fields("pengali_pl_not").Value)
                    chkPL_Not.Value = IIf(IsNull(.Fields("flag_pot_pl_not").Value), 0, .Fields("flag_pot_pl_not").Value)
                    txtSick.Text = IIf(IsNull(.Fields("pengali_sick").Value), 0, .Fields("pengali_sick").Value)
                    chkSick.Value = IIf(IsNull(.Fields("flag_pot_sick").Value), 0, .Fields("flag_pot_sick").Value)
                    txtAlpha.Text = IIf(IsNull(.Fields("pengali_alpha").Value), 0, .Fields("pengali_alpha").Value)
                    chkAlpha.Value = IIf(IsNull(.Fields("flag_pot_alpha").Value), 0, .Fields("flag_pot_alpha").Value)
                End With
            End If
        End If
    ElseIf SSTab1.Tab = 6 Then
        If rsGen.RecordCount > 0 Then
            With rsGen
                v_flag_presence = .Fields("flag_presence").Value
                opt_daily_presence.Value = IIf(v_flag_presence = 0, True, False)
                opt_monthly_presence.Value = IIf(v_flag_presence = 0, False, True)
                txt_presence_allowance.Text = FormatNumber(IIf(IsNull(.Fields("presence_allowance").Value), 0, .Fields("presence_allowance").Value))
                
                v_flag_meal = .Fields("flag_meal").Value
                opt_daily_meal.Value = IIf(v_flag_meal = 0, 1, 0)
                opt_monthly_meal.Value = IIf(v_flag_meal = 0, 0, 1)
                txt_meal_allowance.Text = FormatNumber(IIf(IsNull(.Fields("meal_allowance").Value), 0, .Fields("meal_allowance").Value))
                
                v_flag_transport = .Fields("flag_transport").Value
                opt_daily_transport.Value = IIf(v_flag_transport = 0, True, False)
                opt_monthly_transport.Value = IIf(v_flag_transport = 0, False, True)
                txt_transport_allowance.Text = FormatNumber(IIf(IsNull(.Fields("transport_allowance").Value), 0, .Fields("transport_allowance").Value))
                
                txt_shift2_allowance.Text = FormatNumber(IIf(IsNull(.Fields("shift2_allowance").Value), 0, .Fields("shift2_allowance").Value))
                txt_shift3_allowance.Text = FormatNumber(IIf(IsNull(.Fields("shift3_allowance").Value), 0, .Fields("shift3_allowance").Value))
                
                txt_iuran_koperasi.Text = FormatNumber(IIf(IsNull(.Fields("iuran_koperasi").Value), 0, .Fields("iuran_koperasi").Value))
            End With
        End If
    ElseIf SSTab1.Tab = 7 Then
        If rsPeriode.RecordCount > 0 Then
            With rsPeriode
                vFlagPeriode = .Fields("flag_periode").Value
                optPeriode(0).Value = IIf(vFlagPeriode = 0, True, False)
                optPeriode(1).Value = IIf(vFlagPeriode = 0, False, True)
                
                txt_day_start.Text = .Fields("day_start").Value
                txt_day_end.Text = .Fields("day_end").Value
            End With
        End If
    End If
    
    Exit Sub
    
Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_data()
    int_mode = 2
    Call load_mode
End Sub

Private Sub chkAllDept_Presence_Click()
    If chkAllDept_Presence.Value Then
        TDBCombo_department_presence.Enabled = False
    Else
        TDBCombo_department_presence.Enabled = True
    End If
End Sub

Private Sub chkAllDept_Leave_Click()
    If chkAllDept_Leave.Value Then
        TDBCombo_department_leave.Enabled = False
    Else
        TDBCombo_department_leave.Enabled = True
    End If
End Sub

Private Sub chkPLConvert_Click()
    If chkPLConvert.Value Then
        chkPLConvert.Caption = "YES"
    Else
        chkPLConvert.Caption = "NO"
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

Dim vNo As Integer

    CnG.BeginTrans
    
    If SSTab1.Tab = 0 Then
        SQL = "DELETE FROM m_pref_gen"
        CnG.Execute SQL
        
        SQL = "INSERT INTO m_pref_gen(company_code,late_tolerance," & _
                    "flag_punishment,wh_value,pph21_code,ptkp_code,jamsostek_code,flag_pl," & _
                    "entry_date,entry_user) " & _
                "VALUES " & _
                "('" & TDBCombo_company_gen.Text & "','" & txtLateTolerance.Text & "'," & _
                "'" & chkPunishment.Value & "','" & Val(txt_value_wh.Text) & "'," & _
                "'" & TDBCombo_pph.Text & "','" & TDBCombo_ptkp.Text & "','" & TDBCombo_jstk.Text & "'," & _
                "'" & chkPLConvert.Value & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "DELETE FROM m_pref_ot"
        CnG.Execute SQL
        
        SQL = "INSERT INTO m_pref_ot(flag_ot_method,flag_auto_approve,ot_max_meal,flag_meal,ot_max_trans,flag_trans,auto_ot_sat,auto_ot_wd," _
                & "flag_calc_start,delay_hours,ot_type,ot_sat,ot_wd,on_call_allowance,change_status_hours,entry_date,entry_user) " _
                & "VALUES " _
                & "('" & IIf(optSPL.Value, 0, 1) & "','" & chk_approve.Value & "','" & Val(txtMAXOt(0).Text) & "','" & chk_meal.Value & "'," _
                & "'" & Val(txtMAXOt(1).Text) & "','" & chk_transport.Value & "'," _
                & "'" & Val(txtAutoOt_Sat.Text) & "','" & Val(txtAutoOt_wd.Text) & "'," _
                & "'" & IIf(optAfter.Value, 0, 1) & "','" & Val(txtOTStart.Text) & "'," _
                & "'" & TDBCombo_ot1.Text & "','" & TDBCombo_ot2.Text & "','" & TDBCombo_ot3.Text & "'," _
                & "'" & Val(DropAllComma(txt_on_call_allowance.Text)) & "','" & Val(txt_change_status.Text) & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        If chkAllDept_Presence.Value Then
            SQL = "SELECT department_code FROM m_department WHERE company_code = '" & TDBCombo_company_presence.Text & "'"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                While Not rs.EOF
                    SQL = "DELETE FROM m_pref_presence WHERE company_code = '" & TDBCombo_company_presence.Text & "' and  department_code = '" & rs!DEPARTMENT_CODE & "'"
                    CnG.Execute SQL
                    
                    SQL = "INSERT INTO m_pref_presence(effective_date,company_code,flag_dept,department_code,jhk_value,description,entry_date,entry_user) " _
                            & "VALUES " _
                            & "('" & Format(DTPicker_Presence, "yyyy-MM-dd") & "','" & TDBCombo_company_presence.Text & "',0," _
                            & "'" & rs!DEPARTMENT_CODE & "'," _
                            & "'" & txt_value_jhk.Text & "','" & Trim(txt_description_presence.Text) & "',now(),'" & LOGIN_NAME & "')"
                    CnG.Execute SQL
                    rs.MoveNext
                Wend
            End If
            rs.Close
        Else
            SQL = "INSERT INTO m_pref_presence(effective_date,company_code,flag_dept,department_code,jhk_value,description,entry_date,entry_user) " _
                    & "VALUES " _
                    & "('" & Format(DTPicker_Presence, "yyyy-MM-dd") & "','" & TDBCombo_company_presence.Text & "',0," _
                    & "'" & TDBCombo_department_presence.Text & "'," _
                    & "'" & txt_value_jhk.Text & "','" & Trim(txt_description_presence.Text) & "',now(),'" & LOGIN_NAME & "')"
            CnG.Execute SQL
        End If
    ElseIf SSTab1.Tab = 3 Then
        SQL = "INSERT INTO m_pref_umk(effective_date,umk_value,description,entry_date,entry_user) " _
                & "VALUES " _
                & "('" & Format(DTPicker_umk, "yyyy-MM-dd") & "'," _
                & "'" & Val(DropAllComma(txt_umk_value.Text)) & "','" & Trim(txt_description.Text) & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 4 Then
        If chkAllDept_Leave.Value Then
            SQL = "SELECT department_code FROM m_department WHERE company_code = '" & TDBCombo_company_leave.Text & "'"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                While Not rs.EOF
                    SQL = "DELETE FROM m_pref_leave WHERE company_code = '" & TDBCombo_company_leave.Text & "' and department_code = '" & rs!DEPARTMENT_CODE & "'"
                    CnG.Execute SQL
                    
                    SQL = "INSERT INTO m_pref_leave(effective_date,company_code,flag_dept,department_code,leave_type,description,entry_date,entry_user) " _
                            & "VALUES " _
                            & "('" & Format(DTPicker_Leave, "yyyy-MM-dd") & "','" & TDBCombo_company_leave.Text & "',0," _
                            & "'" & rs!DEPARTMENT_CODE & "'," _
                            & "'" & cboLeave.ListIndex & "','" & Trim(txt_description_leave.Text) & "',now(),'" & LOGIN_NAME & "')"
                    CnG.Execute SQL
                    rs.MoveNext
                Wend
            End If
            rs.Close
        Else
            SQL = "INSERT INTO m_pref_leave(effective_date,company_code,flag_dept,department_code,leave_type,description,entry_date,entry_user) " _
                    & "VALUES " _
                    & "('" & Format(DTPicker_Leave, "yyyy-MM-dd") & "','" & TDBCombo_company_leave.Text & "',0," _
                    & "'" & TDBCombo_department_leave.Text & "'," _
                    & "'" & cboLeave.ListIndex & "','" & Trim(txt_description_leave.Text) & "',now(),'" & LOGIN_NAME & "')"
            CnG.Execute SQL
        End If
    ElseIf SSTab1.Tab = 5 Then
        Dim rs2 As New ADODB.Recordset
                    
        If SSTab2.Tab = 0 Then
            SQL = "SELECT MAX(id_number) id_number FROM m_pref_preall"
            rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs2.RecordCount > 0 Then
                vNo = IIf(IsNull(rs2!id_number), 0, rs2!id_number) + 1
            Else
                vNo = 1
            End If
            rs2.Close
            
            SQL = "INSERT INTO m_pref_preall(id_number,limit_hours,percentage,description,entry_date,entry_user) " _
                    & "VALUES " _
                    & "('" & vNo & "','" & txtLimitHours.Text & "'," _
                    & "'" & txtPercentage.Text & "','" & Trim(txt_description_PreAll.Text) & "',now(),'" & LOGIN_NAME & "')"
            CnG.Execute SQL
        ElseIf SSTab2.Tab = 1 Then
            SQL = "SELECT MAX(id_number) id_number FROM m_pref_lateconvert"
            rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs2.RecordCount > 0 Then
                vNo = IIf(IsNull(rs2!id_number), 0, rs2!id_number) + 1
            Else
                vNo = 1
            End If
            rs2.Close
            
            SQL = "INSERT INTO m_pref_lateconvert(id_number,from_value,to_value,convert_value,description,entry_date,entry_user) " _
                    & "VALUES " _
                    & "('" & vNo & "','" & txtFrom_LateConvert.Text & "','" & txtTo_LateConvert.Text & "'," _
                    & "'" & txtConvert.Text & "','" & Trim(txt_description_LateConvert.Text) & "',now(),'" & LOGIN_NAME & "')"
            CnG.Execute SQL
        ElseIf SSTab2.Tab = 2 Then
            SQL = "DELETE FROM m_pref_pen_gen"
            CnG.Execute SQL
            
            SQL = "INSERT INTO m_pref_pen_gen(pengali_pl,flag_pot_pl,pengali_pl_not,flag_pot_pl_not," _
                    & "pengali_sick,flag_pot_sick,pengali_alpha,flag_pot_alpha,entry_date,entry_user) " _
                    & "VALUES " _
                    & "('" & txtPL.Text & "','" & chkPL.Value & "','" & txtPL_Not.Text & "','" & chkPL_Not.Value & "'," _
                    & "'" & txtSick.Text & "','" & chkSick.Value & "'," _
                    & "'" & txtAlpha.Text & "','" & chkAlpha.Value & "',now(),'" & LOGIN_NAME & "')"
            CnG.Execute SQL
        End If
    ElseIf SSTab1.Tab = 6 Then
        SQL = "DELETE FROM m_pref_allow"
        CnG.Execute SQL
        
        SQL = "INSERT INTO m_pref_allow (flag_presence,presence_allowance," & _
            "flag_transport,transport_allowance,flag_meal,meal_allowance,shift2_allowance,shift3_allowance,iuran_koperasi,entry_date,entry_user) " & _
            "VALUES " & _
            "('" & IIf(opt_daily_presence.Value, 0, 1) & "'," & Val(DropAllComma(txt_presence_allowance.Text)) & "," & _
            "'" & IIf(opt_daily_transport.Value, 0, 1) & "'," & Val(DropAllComma(txt_transport_allowance.Text)) & "," & _
            "'" & IIf(opt_daily_meal.Value, 0, 1) & "'," & Val(DropAllComma(txt_meal_allowance.Text)) & "," & _
            "" & Val(DropAllComma(txt_shift2_allowance.Text)) & ",'" & Val(DropAllComma(txt_shift3_allowance.Text)) & "'," & _
            "" & Val(DropAllComma(txt_iuran_koperasi.Text)) & ",now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
        
        If chkAll.Value = 1 Then
            Dim vTgl As String
            Dim i As Integer
            
            i = MsgBox("Apply To All Will Update All Last Data Salary Employee..." & Chr(13) & _
                        "Are You Sure To Update This Data?", vbYesNo + vbQuestion, headerMSG)
            If Not i = vbYes Then
                CnG.RollbackTrans
                Exit Sub
            End If
        
            SQL = "SELECT * FROM m_employee WHERE flag_active <> 0"
            rs.Open SQL, CnG, adOpenForwardOnly
            
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                While Not rs.EOF
                    SQL = "SELECT date(salary_date) from m_salary_standard " & _
                            "WHERE employee_code = '" & rs!employee_code & "' " & _
                            "AND date(salary_date) <= date(now()) ORDER BY salary_date DESC LIMIT 1"
                    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rscari.RecordCount > 0 Then
                        vTgl = Format(rscari.Fields(0).Value, "yyyy-MM-dd")
                    Else
                        vTgl = Format(Now, "yyyy-MM-dd")
                    End If
                    rscari.Close
                            
                    SQL = "UPDATE m_salary_standard SET flag_presence = '" & IIf(opt_daily_presence.Value, 0, 1) & "'," & _
                            "presence_allowance = '" & Val(DropAllComma(txt_presence_allowance.Text)) & "'," & _
                            "flag_transport = '" & IIf(opt_daily_transport.Value, 0, 1) & "'," & _
                            "transport_allowance = '" & Val(DropAllComma(txt_transport_allowance.Text)) & "'," & _
                            "flag_meal = '" & IIf(opt_daily_meal.Value, 0, 1) & "'," & _
                            "meal_allowance = '" & Val(DropAllComma(txt_meal_allowance.Text)) & "'," & _
                            "shift2_allowance = '" & Val(DropAllComma(txt_shift2_allowance.Text)) & "'," & _
                            "shift2_allowance = '" & Val(DropAllComma(txt_shift3_allowance.Text)) & "' " & _
                          "WHERE employee_code = '" & rs!employee_code & "' " & _
                            "AND date(salary_date) = '" & vTgl & "' "
                    CnG.Execute SQL
                    
                    SQL = "UPDATE m_salary_standard SET coop_value = '" & Val(DropAllComma(txt_iuran_koperasi.Text)) & "' " & _
                          "WHERE employee_code = '" & rs!employee_code & "' " & _
                            "AND date(salary_date) = '" & vTgl & "' " & _
                            "AND flag_coop = 1"
                    CnG.Execute SQL
                rs.MoveNext
                Wend
            End If
            rs.Close
        End If
        MsgBox "Save Succesfully...", vbInformation, headerMSG
    ElseIf SSTab1.Tab = 7 Then
        SQL = "DELETE FROM m_pref_periode"
        CnG.Execute SQL
        
        SQL = "INSERT INTO m_pref_periode(flag_periode,day_start,day_end) " & _
                "VALUES " & _
                "(" & IIf(optPeriode(0).Value, 0, 1) & "," & txt_day_start.Text & "," & txt_day_end.Text & ")"
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
    
    If SSTab1.Tab = 2 Then
        SQL = "UPDATE m_pref_presence SET jhk_value = '" & txt_value_jhk.Text & "'," _
                & "description = '" & Trim(txt_description_presence.Text) & "', " _
                & "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " _
                & "WHERE company_code = '" & TDBCombo_company_presence.Text & "' " _
                    & "and department_code = '" & TDBCombo_department_presence.Text & "' " _
                    & "and date(effective_date) = '" & Format(DTPicker_Presence.Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 3 Then
        SQL = "UPDATE m_pref_umk SET umk_value = '" & Val(DropAllComma(txt_umk_value)) & "'," _
                & "description = '" & Trim(txt_description.Text) & "', " _
                & "edit_date = now(), " _
                & "edit_user = '" & LOGIN_NAME & "' " _
                & "WHERE date(effective_date) = '" & Format(DTPicker_umk.Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 4 Then
        SQL = "UPDATE m_pref_leave SET leave_type = '" & cboLeave.ListIndex & "'," _
                & "description = '" & Trim(txt_description_leave.Text) & "', " _
                & "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " _
                & "WHERE company_code = '" & TDBCombo_company_leave.Text & "' " _
                    & "and department_code = '" & TDBCombo_department_leave.Text & "' " _
                    & "and date(effective_date) = '" & Format(DTPicker_Leave.Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 5 Then
        If SSTab2.Tab = 0 Then
            SQL = "UPDATE m_pref_preall SET limit_hours = '" & txtLimitHours.Text & "'," _
                    & "percentage = '" & txtPercentage.Text & "', " _
                    & "description = '" & Trim(txt_description_PreAll.Text) & "', " _
                    & "edit_date = now(), " _
                    & "edit_user = '" & LOGIN_NAME & "' " _
                    & "WHERE id_number = '" & vIdNumber & "'"
            CnG.Execute SQL
        ElseIf SSTab2.Tab = 1 Then
            SQL = "UPDATE m_pref_lateconvert SET from_value = '" & txtFrom_LateConvert.Text & "'," _
                    & "to_value = '" & txtTo_LateConvert.Text & "', " _
                    & "convert_value = '" & txtConvert.Text & "', " _
                    & "description = '" & Trim(txt_description_LateConvert.Text) & "', " _
                    & "edit_date = now(), " _
                    & "edit_user = '" & LOGIN_NAME & "' " _
                    & "WHERE id_number = '" & vIdNumber & "'"
            CnG.Execute SQL
        End If
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
    If SSTab1.Tab = 2 Then
        cmdNew_Presence.Enabled = a And blnUser_Add
        cmdSave_Presence.Enabled = b
        cmdEdit_Presence.Enabled = c And blnUser_Edit
        cmdDelete_Presence.Enabled = d And blnUser_Delete
        cmdCancel_Presence.Enabled = e
    ElseIf SSTab1.Tab = 3 Then
        cmdNew_UMK.Enabled = a And blnUser_Add
        cmdSave_UMK.Enabled = b
        cmdEdit_UMK.Enabled = c And blnUser_Edit
        cmdDelete_UMK.Enabled = d And blnUser_Delete
        cmdCancel_UMK.Enabled = e
    ElseIf SSTab1.Tab = 4 Then
        cmdNew_Leave.Enabled = a And blnUser_Add
        cmdSave_Leave.Enabled = b
        cmdEdit_Leave.Enabled = c And blnUser_Edit
        cmdDelete_Leave.Enabled = d And blnUser_Delete
        cmdCancel_Leave.Enabled = e
    ElseIf SSTab1.Tab = 5 Then
        If SSTab2.Tab = 0 Then
            cmdNew_PreAll.Enabled = a And blnUser_Add
            cmdSave_PreAll.Enabled = b
            cmdEdit_PreAll.Enabled = c And blnUser_Edit
            cmdDelete_PreAll.Enabled = d And blnUser_Delete
            cmdCancel_PreAll.Enabled = e
        ElseIf SSTab2.Tab = 1 Then
            cmdNew_LateConvert.Enabled = a And blnUser_Add
            cmdSave_LateConvert.Enabled = b
            cmdEdit_LateConvert.Enabled = c And blnUser_Edit
            cmdDelete_LateConvert.Enabled = d And blnUser_Delete
            cmdCancel_LateConvert.Enabled = e
        End If
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
        
        If SSTab1.Tab = 2 Then
            fra_entry_jhk.Visible = True
            TDBCombo_company_presence.Enabled = True
            chkAllDept_Presence.Enabled = True
            TDBCombo_department_presence.Enabled = True
            TDBGrid_Presence.Enabled = False
            Call set_new_data
            
            If TDBCombo_company_presence.Enabled = True Then
                TDBCombo_company_presence.SetFocus
            End If
        ElseIf SSTab1.Tab = 3 Then
            fra_entry.Visible = True
            DTPicker_umk.Enabled = True
            TDBGrid_UMK.Enabled = False
            Call set_new_data
            
            If DTPicker_umk.Enabled = True Then
                DTPicker_umk.SetFocus
            End If
        ElseIf SSTab1.Tab = 4 Then
            fra_entry_leave.Visible = True
            TDBCombo_company_leave.Enabled = True
            chkAllDept_Leave.Enabled = True
            TDBCombo_department_leave.Enabled = True
            TDBGrid_Leave.Enabled = False
            Call set_new_data
            
            If TDBCombo_company_leave.Enabled = True Then
                TDBCombo_company_leave.SetFocus
            End If
        ElseIf SSTab1.Tab = 5 Then
            If SSTab2.Tab = 0 Then
                fra_entry_PreAll.Visible = True
                TDBGrid_PreAll.Enabled = False
                Call set_new_data
                
                txtLimitHours.SetFocus
            ElseIf SSTab2.Tab = 1 Then
                fra_entry_LateConvert.Visible = True
                TDBGrid_LateConvert.Enabled = False
                Call set_new_data
                
                txtFrom_LateConvert.SetFocus
            End If
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        
        If SSTab1.Tab = 2 Then
            fra_entry_jhk.Visible = False
            TDBGrid_Presence.Enabled = True
        ElseIf SSTab1.Tab = 3 Then
            fra_entry.Visible = False
            TDBGrid_UMK.Enabled = True
        ElseIf SSTab1.Tab = 4 Then
            fra_entry_leave.Visible = False
            TDBGrid_Leave.Enabled = True
        ElseIf SSTab1.Tab = 5 Then
            If SSTab2.Tab = 0 Then
                fra_entry_PreAll.Visible = False
                TDBGrid_PreAll.Enabled = True
            ElseIf SSTab2.Tab = 1 Then
                fra_entry_LateConvert.Visible = False
                TDBGrid_LateConvert.Enabled = True
            End If
        End If
        
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        If SSTab1.Tab = 2 Then
            DTPicker_Presence.Enabled = False
            TDBCombo_company_presence.Enabled = False
            chkAllDept_Presence.Enabled = False
            TDBCombo_department_presence.Enabled = False
            fra_entry_jhk.Visible = True
            TDBGrid_Presence.Enabled = False
        ElseIf SSTab1.Tab = 3 Then
            DTPicker_umk.Enabled = False
            fra_entry.Visible = True
            TDBGrid_UMK.Enabled = False
        ElseIf SSTab1.Tab = 4 Then
            DTPicker_Leave.Enabled = False
            TDBCombo_company_leave.Enabled = False
            chkAllDept_Leave.Enabled = False
            TDBCombo_department_leave.Enabled = False
            fra_entry_leave.Visible = True
            TDBGrid_Leave.Enabled = False
        ElseIf SSTab1.Tab = 5 Then
            If SSTab2.Tab = 0 Then
                fra_entry_PreAll.Visible = True
                TDBGrid_PreAll.Enabled = False
            ElseIf SSTab2.Tab = 1 Then
                fra_entry_LateConvert.Visible = True
                TDBGrid_LateConvert.Enabled = False
            End If
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

Private Sub cmdSave_Allow_Click()
    Call insert_new_data
End Sub

Private Sub cmdSave_Att_Click()
    Call insert_new_data
    
    MsgBox "Save Succesfully...", vbInformation, headerMSG
End Sub

Private Sub cmdSave_Gen_Click()
    Call insert_new_data
    
    MsgBox "Save Succesfully...", vbInformation, headerMSG
End Sub

Private Sub cmdSave_OT_Click()
    If optSystem.Value Then
        If TDBCombo_ot1.Text = "" Then
            MsgBox "OT Type Is Not Selected...", vbExclamation, headerMSG
            Exit Sub
        End If
        
        If TDBCombo_ot3.Text = "" Then
            MsgBox "OT Work Day Is Not Selected...", vbExclamation, headerMSG
            Exit Sub
        End If
        
        If TDBCombo_ot2.Text = "" Then
            MsgBox "OT Saturday Is Not Selected...", vbExclamation, headerMSG
            Exit Sub
        End If
    End If
     
    Call insert_new_data
    
    MsgBox "Save Succesfully...", vbInformation, headerMSG
End Sub

Private Sub cmdSave_Periode_Click()
    Call insert_new_data
    
    MsgBox "Save Succesfully...", vbInformation, headerMSG
End Sub

Private Sub Form_Load()
    cboLeave.ListIndex = 0
    oClause = ""
    
'    Call load_data_company
'    Call load_data_pph21
'    Call load_data_ptkp
'    Call load_data_jstk
'    txtLateTolerance.SetFocus
    
    SSTab1.Tab = 0
    Call load_data_user_access(Me)
    Call load_data
    
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    SSTab1.TabVisible(4) = False
    
'    If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Then
'        Call load_data
'        Exit Sub
'    Else
'        Call load_data
'
'        int_mode = 0
'        Call load_mode
'        Exit Sub
'    End If
End Sub

Private Sub clear_filter()
    If SSTab1.Tab = 2 Then
        For Each Col In TDBGrid_Presence.Columns
            Col.FilterText = ""
        Next Col
        rsJHK.Filter = adFilterNone
    ElseIf SSTab1.Tab = 3 Then
        For Each Col In TDBGrid_UMK.Columns
            Col.FilterText = ""
        Next Col
        rsUMK.Filter = adFilterNone
    ElseIf SSTab1.Tab = 4 Then
        For Each Col In TDBGrid_Leave.Columns
            Col.FilterText = ""
        Next Col
        rsLeave.Filter = adFilterNone
    ElseIf SSTab1.Tab = 5 Then
        If SSTab2.Tab = 0 Then
            For Each Col In TDBGrid_PreAll.Columns
                Col.FilterText = ""
            Next Col
            rsPreAll.Filter = adFilterNone
        ElseIf SSTab2.Tab = 1 Then
            For Each Col In TDBGrid_LateConvert.Columns
                Col.FilterText = ""
            Next Col
            rsLateConvert.Filter = adFilterNone
        End If
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
    
    If SSTab1.Tab = 2 Then
        Set Cols = TDBGrid_Presence.Columns
        i = TDBGrid_Presence.Col
        TDBGrid_Presence.HoldFields
        
        rsJHK.Filter = getFilter()
        TDBGrid_Presence.Col = i
        TDBGrid_Presence.EditActive = True
        
        TDBGrid_Presence.SelStart = Len(TDBGrid_Presence.Columns(i).FilterText)
        If TDBGrid_Presence.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Presence.Col = i
        End If
    ElseIf SSTab1.Tab = 3 Then
        Set Cols = TDBGrid_UMK.Columns
        i = TDBGrid_UMK.Col
        TDBGrid_UMK.HoldFields
        
        rsUMK.Filter = getFilter()
        TDBGrid_UMK.Col = i
        TDBGrid_UMK.EditActive = True
        
        TDBGrid_UMK.SelStart = Len(TDBGrid_UMK.Columns(i).FilterText)
        If TDBGrid_UMK.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_UMK.Col = i
        End If
    ElseIf SSTab1.Tab = 4 Then
        Set Cols = TDBGrid_Leave.Columns
        i = TDBGrid_Leave.Col
        TDBGrid_Leave.HoldFields
        
        rsLeave.Filter = getFilter()
        TDBGrid_Leave.Col = i
        TDBGrid_Leave.EditActive = True
        
        TDBGrid_Leave.SelStart = Len(TDBGrid_Leave.Columns(i).FilterText)
        If TDBGrid_Leave.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Leave.Col = i
        End If
    ElseIf SSTab1.Tab = 5 Then
        If SSTab2.Tab = 0 Then
            Set Cols = TDBGrid_PreAll.Columns
            i = TDBGrid_PreAll.Col
            TDBGrid_PreAll.HoldFields
            
            rsPreAll.Filter = getFilter()
            TDBGrid_PreAll.Col = i
            TDBGrid_PreAll.EditActive = True
            
            TDBGrid_PreAll.SelStart = Len(TDBGrid_PreAll.Columns(i).FilterText)
            If TDBGrid_PreAll.ApproxCount < 1 Then
                Call clear_filter
                TDBGrid_PreAll.Col = i
            End If
        ElseIf SSTab2.Tab = 1 Then
            Set Cols = TDBGrid_LateConvert.Columns
            i = TDBGrid_PreAll.Col
            TDBGrid_LateConvert.HoldFields
            
            rsLateConvert.Filter = getFilter()
            TDBGrid_LateConvert.Col = i
            TDBGrid_LateConvert.EditActive = True
            
            TDBGrid_LateConvert.SelStart = Len(TDBGrid_LateConvert.Columns(i).FilterText)
            If TDBGrid_LateConvert.ApproxCount < 1 Then
                Call clear_filter
                TDBGrid_LateConvert.Col = i
            End If
        End If
    End If

    Exit Sub

Err:
MsgBox "No Data found in this column " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data()
    If SSTab1.Tab = 0 Then
        If rsAtt.State Then rsAtt.Close
        SQL = "select a.*, b.company_name, c.pph21_name, d.ptkp_name, e.jamsostek_name " & _
                "from m_pref_gen a join m_company b on a.company_code = b.company_code " & _
                    "join m_pph21 c on a.pph21_code = c.pph21_code " & _
                    "join m_ptkp d on a.ptkp_code = d.ptkp_code " & _
                    "join m_jamsostek e on a.jamsostek_code = e.jamsostek_code"
        rsAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        Call set_edit_data
    ElseIf SSTab1.Tab = 1 Then
        If rsot.State Then rsot.Close
        SQL = "SELECT a.*, (SELECT ot_name FROM m_ot WHERE ot_code = a.ot_type) ot_name1," & _
                "(SELECT ot_name FROM m_ot WHERE ot_code = a.ot_sat) ot_name2," & _
                "(SELECT ot_name FROM m_ot WHERE ot_code = a.ot_wd) ot_name3 " & _
              "FROM m_pref_ot a"
        rsot.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        Call set_edit_data
    ElseIf SSTab1.Tab = 2 Then
        If rsJHK.State Then rsJHK.Close
        SQL = "select a.*,b.company_name,c.department_name from m_pref_presence a join m_company b on a.company_code = b.company_code " & _
                "LEFT JOIN m_department c on a.company_code = c.company_code and a.department_code = c.department_code " & oClause
        rsJHK.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Presence.DataSource = rsJHK
    ElseIf SSTab1.Tab = 3 Then
        If rsUMK.State Then rsUMK.Close
        SQL = "select * from m_pref_umk " & oClause
        rsUMK.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_UMK.DataSource = rsUMK
    ElseIf SSTab1.Tab = 4 Then
        If rsLeave.State Then rsLeave.Close
        SQL = "select a.*,CASE WHEN leave_type = 0 then 'PERIODE KERJA' else 'TAHUN' end type,b.company_name,c.department_name " & _
                "from m_pref_leave a join m_company b on a.company_code = b.company_code " & _
                "LEFT JOIN m_department c on a.company_code = c.company_code and a.department_code = c.department_code " & oClause
        rsLeave.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Leave.DataSource = rsLeave
    ElseIf SSTab1.Tab = 5 Then
        If SSTab2.Tab = 0 Then
            If rsPreAll.State Then rsPreAll.Close
            SQL = "select * from m_pref_preall " & oClause
            rsPreAll.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            TDBGrid_PreAll.DataSource = rsPreAll
        ElseIf SSTab2.Tab = 1 Then
            If rsLateConvert.State Then rsLateConvert.Close
            SQL = "select * from m_pref_lateconvert " & oClause
            rsLateConvert.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            TDBGrid_LateConvert.DataSource = rsLateConvert
        ElseIf SSTab2.Tab = 2 Then
            If rsPenaltyGeneral.State Then rsPenaltyGeneral.Close
            SQL = "select * from m_pref_pen_gen "
            rsPenaltyGeneral.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            Call set_edit_data
        End If
    ElseIf SSTab1.Tab = 6 Then
        If rsGen.State Then rsGen.Close
        SQL = "select * " & _
                "from m_pref_allow "
        rsGen.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        Call set_edit_data
    ElseIf SSTab1.Tab = 7 Then
        If rsPeriode.State Then rsPeriode.Close
        SQL = "select * " & _
                "from m_pref_periode "
        rsPeriode.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        Call set_edit_data
    End If
End Sub

Private Sub optAfter_Click()
    txtOTStart.Text = ""
    txtOTStart.Enabled = False
End Sub

Private Sub optDelay_Click()
    txtOTStart.Enabled = True
End Sub

Private Sub optPeriode_Click(Index As Integer)
    If Index = 0 Then
        fraPeriode.Visible = False
    Else
        fraPeriode.Visible = True
    End If
End Sub

Private Sub optSPL_Click()
    fraSystem.Visible = False
    chk_approve.Visible = True
End Sub

Private Sub optSystem_Click()
    fraSystem.Visible = True
    chk_approve.Visible = False
    Call load_data_ot_tdb
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Call load_data_user_access(Me)

    oClause = ""
    Call load_data
    
    If SSTab1.Tab = 6 Then
        If LOGIN_LEVEL = 0 Then
            cmdSave_Allow.Enabled = False
        Else
            cmdSave_Allow.Enabled = True
        End If
    ElseIf SSTab1.Tab = 7 Then
        If LOGIN_LEVEL = 0 Then
            cmdSave_Periode.Enabled = False
        Else
            cmdSave_Periode.Enabled = True
        End If
    End If
    
    If SSTab1.Tab = 0 Then
        Call load_data_company
        Call load_data_pph21
        Call load_data_ptkp
        Call load_data_jstk

        Call load_data
        
        cmdSave_Att.Enabled = blnUser_Add
    ElseIf SSTab1.Tab = 2 Or SSTab1.Tab = 4 Then
        Call load_data_company
        
        int_mode = 0
        Call load_mode
    ElseIf SSTab1.Tab = 3 Then
        int_mode = 0
        Call load_mode
    ElseIf SSTab1.Tab = 1 Then
        If vSetData = 1 Then
            Call load_data_ot_tdb
        End If
        
        cmdSave_OT.Enabled = blnUser_Add
    ElseIf SSTab1.Tab = 5 Then
        SSTab2.Tab = 0
        Call load_data
        
        If SSTab1.Tab = 2 Then
            txtPL.SetFocus
            
            cmdSave_Gen.Enabled = blnUser_Add
        Else
            int_mode = 0
            Call load_mode
        End If
    End If
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
    oClause = ""
    Call load_data
    
    If SSTab2.Tab <> 2 Then
        int_mode = 0
        Call load_mode
    Else
        SQL = "SELECT wh_value FROM m_pref_gen"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            lblPL.Caption = rs!wh_value
            lblPL_Not.Caption = rs!wh_value
        Else
            lblPL.Caption = "173"
            lblPL_Not.Caption = "173"
        End If
        rs.Close
    End If
End Sub

Private Sub TDBCombo_company_gen_ItemChange()
    If TDBCombo_company_gen.ApproxCount > 0 Then
        TDBCombo_company_gen.Text = TDBCombo_company_gen.Columns("company_code").Value
        txt_company_name_gen.Text = TDBCombo_company_gen.Columns("company_name").Value
    End If
End Sub

Private Sub TDBCombo_pph_ItemChange()
    If TDBCombo_pph.ApproxCount > 0 Then
        TDBCombo_pph.Text = TDBCombo_pph.Columns("pph21_code").Value
        txt_pph21_name = TDBCombo_pph.Columns("pph21_name").Value
        
    End If
End Sub

Private Sub TDBCombo_ptkp_ItemChange()
    If TDBCombo_ptkp.ApproxCount > 0 Then
        TDBCombo_ptkp.Text = TDBCombo_ptkp.Columns("ptkp_code").Value
        txt_ptkp_name = TDBCombo_ptkp.Columns("ptkp_name").Value
        
    End If
End Sub

Private Sub TDBCombo_jstk_ItemChange()
    If TDBCombo_jstk.ApproxCount > 0 Then
        TDBCombo_jstk.Text = TDBCombo_jstk.Columns("jamsostek_code").Value
        txt_jstk_name = TDBCombo_jstk.Columns("jamsostek_name").Value
        
    End If
End Sub

Private Sub TDBCombo_company_presence_ItemChange()
    If TDBCombo_company_presence.ApproxCount > 0 Then
        TDBCombo_company_presence.Text = TDBCombo_company_presence.Columns("company_code").Value
        txt_company_name_presence.Text = TDBCombo_company_presence.Columns("company_name").Value

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

Private Sub TDBCombo_department_presence_itemChange()
    If TDBCombo_department_presence.ApproxCount > 0 Then
        TDBCombo_department_presence.Text = TDBCombo_department_presence.Columns("department_code").Value
        txt_department_name_presence.Text = TDBCombo_department_presence.Columns("department_name").Value
    End If
End Sub

Private Sub TDBCombo_department_leave_itemChange()
    If TDBCombo_department_leave.ApproxCount > 0 Then
        TDBCombo_department_leave.Text = TDBCombo_department_leave.Columns("department_code").Value
        txt_department_name_leave.Text = TDBCombo_department_leave.Columns("department_name").Value
    End If
End Sub

Private Sub TDBCombo_ot1_ItemChange()
    If TDBCombo_ot1.ApproxCount > 0 Then
        TDBCombo_ot1.Text = TDBCombo_ot1.Columns("ot_code").Value
        txt_ot_name1.Text = TDBCombo_ot1.Columns("ot_name").Value
    End If
End Sub

Private Sub TDBCombo_ot2_ItemChange()
    If TDBCombo_ot2.ApproxCount > 0 Then
        TDBCombo_ot2.Text = TDBCombo_ot2.Columns("ot_code").Value
        txt_ot_name2.Text = TDBCombo_ot2.Columns("ot_name").Value
    End If
End Sub

Private Sub TDBCombo_ot3_ItemChange()
    If TDBCombo_ot3.ApproxCount > 0 Then
        TDBCombo_ot3.Text = TDBCombo_ot3.Columns("ot_code").Value
        txt_ot_name3.Text = TDBCombo_ot3.Columns("ot_name").Value
    End If
End Sub


Public Sub load_data_ot_tdb()
    TDBCombo_ot1.Text = "": txt_ot_name1.Text = ""
    TDBCombo_ot2.Text = "": txt_ot_name2.Text = ""
    TDBCombo_ot3.Text = "": txt_ot_name3.Text = ""
    
    If rsOT_TDB.State Then rsOT_TDB.Close
    SQL = "select * from m_ot order by ot_code"
    rsOT_TDB.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_ot1.RowSource = rsOT_TDB
    TDBCombo_ot2.RowSource = rsOT_TDB
    TDBCombo_ot3.RowSource = rsOT_TDB
End Sub
Public Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
    If SSTab1.Tab = 0 Then
        TDBCombo_company_gen.Text = "": txt_company_name_gen = ""
        TDBCombo_company_gen.RowSource = rsCompany
    ElseIf SSTab1.Tab = 2 Then
        TDBCombo_company_presence.Text = "": txt_company_name_presence = ""
        TDBCombo_company_presence.RowSource = rsCompany
    ElseIf SSTab1.Tab = 4 Then
        TDBCombo_company_leave.Text = "": txt_company_name_leave = ""
        TDBCombo_company_leave.RowSource = rsCompany
    End If
End Sub

Public Sub load_data_division()
    If SSTab1.Tab = 2 Then
        TDBCombo_department_presence.Text = "": txt_department_name_presence.Text = ""
        
        If rsDivision.State Then rsDivision.Close
        SQL = "select * from m_department where company_code = '" & TDBCombo_company_presence.Text & "' order by company_code"
        rsDivision.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBCombo_department_presence.RowSource = rsDivision
    ElseIf SSTab1.Tab = 4 Then
        TDBCombo_department_leave.Text = "": txt_department_name_leave.Text = ""
        
        If rsDivision.State Then rsDivision.Close
        SQL = "select * from m_department where company_code = '" & TDBCombo_company_leave.Text & "' order by company_code"
        rsDivision.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBCombo_department_leave.RowSource = rsDivision
    End If
End Sub

Private Sub load_data_pph21()
    If rsPPh.State Then rsPPh.Close
    SQL = "select * from m_pph21 order by pph21_code"
    rsPPh.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_pph.RowSource = rsPPh
End Sub

Private Sub load_data_ptkp()
    If rsPTKP.State Then rsPTKP.Close
    SQL = "select * from m_ptkp order by ptkp_code"
    rsPTKP.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_ptkp.RowSource = rsPTKP
End Sub

Private Sub load_data_jstk()
    If rsJSTK.State Then rsJSTK.Close
    SQL = "select * from m_jamsostek order by jamsostek_code"
    rsJSTK.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_jstk.RowSource = rsJSTK
End Sub

Private Sub txt_umk_value_Validate(Cancel As Boolean)
    If Not Trim(txt_umk_value) = "" Then
        txt_umk_value = FormatNumber(DropAllComma(txt_umk_value))
    End If
End Sub

Private Sub tdbGrid_Presence_FilterChange()
    Call filter_change
End Sub

Private Sub tdbGrid_UMK_FilterChange()
    Call filter_change
End Sub

Private Sub tdbGrid_Leave_FilterChange()
    Call filter_change
End Sub

Private Sub tdbGrid_PreAll_FilterChange()
    Call filter_change
End Sub

Private Sub tdbGrid_LateConvert_FilterChange()
    Call filter_change
End Sub


Private Sub cmdNew_Presence_Click()
    Call new_data
End Sub

Private Sub cmdSave_Presence_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Presence_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Presence_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Presence_Click()
    Call cancel_data
End Sub

Private Sub cmdNew_UMK_Click()
    Call new_data
End Sub

Private Sub cmdSave_UMK_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_UMK_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_UMK_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_UMK_Click()
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

Private Sub cmdNew_PreAll_Click()
    Call new_data
End Sub

Private Sub cmdSave_PreAll_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_PreAll_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_PreAll_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_PreAll_Click()
    Call cancel_data
End Sub

Private Sub cmdNew_LateConvert_Click()
    Call new_data
End Sub

Private Sub cmdSave_LateConvert_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_LateConvert_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_LateConvert_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_LateConvert_Click()
    Call cancel_data
End Sub

Private Sub TDBGrid_Presence_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Presence.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Presence.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Presence.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Presence.Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub TDBGrid_UMK_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_UMK.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_UMK.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_UMK.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_UMK.Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub TDBGrid_Leave_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Leave.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Leave.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Leave.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Leave.Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub TDBGrid_PreAll_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_PreAll.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_PreAll.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_PreAll.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_PreAll.Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub TDBGrid_LateConvert_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_LateConvert.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_LateConvert.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_LateConvert.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_LateConvert.Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub txt_presence_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_presence_allowance) = "" Then
        txt_presence_allowance = FormatNumber(DropAllComma(txt_presence_allowance))
    End If
End Sub

Private Sub txt_transport_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_transport_allowance) = "" Then
        txt_transport_allowance = FormatNumber(DropAllComma(txt_transport_allowance))
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

Private Sub txt_iuran_koperasi_Validate(Cancel As Boolean)
    If Not Trim(txt_iuran_koperasi) = "" Then
        txt_iuran_koperasi = FormatNumber(DropAllComma(txt_iuran_koperasi))
    End If
End Sub

Private Sub txt_on_call_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_on_call_allowance) = "" Then
        txt_on_call_allowance = FormatNumber(DropAllComma(txt_on_call_allowance))
    End If
End Sub
