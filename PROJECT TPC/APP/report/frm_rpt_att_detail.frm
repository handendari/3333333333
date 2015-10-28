VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_rpt_detail_attendance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORT - ATTENDANCE"
   ClientHeight    =   7350
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   10530
   Icon            =   "frm_rpt_att_detail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   3465
      Left            =   4680
      TabIndex        =   51
      Top             =   3480
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
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   5940
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   180
         Top             =   360
      End
      Begin prj_tpc.vbButton cmdExit 
         Height          =   705
         Left            =   8190
         TabIndex        =   43
         Top             =   300
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
         MICON           =   "frm_rpt_att_detail.frx":058A
         PICN            =   "frm_rpt_att_detail.frx":05A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdPrint 
         Height          =   705
         Left            =   7170
         TabIndex        =   44
         Top             =   300
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
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
         MICON           =   "frm_rpt_att_detail.frx":1638
         PICN            =   "frm_rpt_att_detail.frx":1654
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
      Left            =   3150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   46
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3150
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   840
      Width           =   3975
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1350
      OleObjectBlob   =   "frm_rpt_att_detail.frx":26E6
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_department 
      Height          =   375
      Left            =   1350
      OleObjectBlob   =   "frm_rpt_att_detail.frx":46A4
      TabIndex        =   47
      Top             =   1200
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4155
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7329
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "DETAIL"
      TabPicture(0)   =   "frm_rpt_att_detail.frx":6665
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SUMMARY"
      TabPicture(1)   =   "frm_rpt_att_detail.frx":6681
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CALCULATION"
      TabPicture(2)   =   "frm_rpt_att_detail.frx":669D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "MAN WORKING"
      TabPicture(3)   =   "frm_rpt_att_detail.frx":66B9
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   2865
         Left            =   -74160
         TabIndex        =   55
         Top             =   690
         Width           =   8415
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   135
            Left            =   1800
            TabIndex        =   76
            Top             =   2550
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   238
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
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
            Height          =   315
            Left            =   3600
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   390
            Width           =   3975
         End
         Begin VB.Frame Frame11 
            Caption         =   "TYPE"
            Height          =   525
            Left            =   1800
            TabIndex        =   70
            Top             =   1140
            Width           =   4125
            Begin VB.OptionButton opt_periode 
               Caption         =   "SUMMARY MAN HOUR"
               Height          =   255
               Left            =   1920
               TabIndex        =   72
               Top             =   210
               Width           =   2115
            End
            Begin VB.OptionButton opt_monthly 
               Caption         =   "SUMMARY ATT"
               Height          =   255
               Left            =   180
               TabIndex        =   71
               Top             =   210
               Width           =   1725
            End
         End
         Begin VB.ComboBox cbo_employee 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":66D5
            Left            =   1800
            List            =   "frm_rpt_att_detail.frx":66DF
            TabIndex        =   57
            Text            =   "..."
            Top             =   780
            Width           =   1695
         End
         Begin VB.TextBox txt_employee_code 
            Height          =   315
            Left            =   7830
            TabIndex        =   56
            Top             =   780
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Frame fra_monthly 
            Height          =   585
            Left            =   1800
            TabIndex        =   68
            Top             =   1590
            Width           =   1875
            Begin MSComCtl2.DTPicker DTPicker_monthly 
               Height          =   300
               Left            =   90
               TabIndex        =   69
               Top             =   180
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "MM-yyyy"
               Format          =   124125187
               UpDown          =   -1  'True
               CurrentDate     =   39278
            End
         End
         Begin VB.Frame fra_periode 
            Height          =   585
            Left            =   1800
            TabIndex        =   64
            Top             =   1590
            Width           =   4125
            Begin MSComCtl2.DTPicker DTPicker_periode_from 
               Height          =   300
               Left            =   150
               TabIndex        =   65
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   124125187
               CurrentDate     =   39278
            End
            Begin MSComCtl2.DTPicker DTPicker_periode_to 
               Height          =   300
               Left            =   2280
               TabIndex        =   66
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   124125187
               CurrentDate     =   39278
            End
            Begin VB.Label Label11 
               Caption         =   "TO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1950
               TabIndex        =   67
               Top             =   210
               Width           =   285
            End
         End
         Begin VB.Frame fra_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   58
            Top             =   540
            Width           =   4575
            Begin VB.TextBox txt_nik 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               Height          =   315
               Left            =   0
               MaxLength       =   50
               TabIndex        =   60
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txt_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   59
               Top             =   240
               Width           =   2415
            End
            Begin prj_tpc.vbButton cmd_browse 
               Height          =   315
               Left            =   1410
               TabIndex        =   61
               Top             =   240
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               BTYPE           =   14
               TX              =   "..."
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
               MICON           =   "frm_rpt_att_detail.frx":66F0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_division 
            Height          =   375
            Left            =   1800
            OleObjectBlob   =   "frm_rpt_att_detail.frx":670C
            TabIndex        =   74
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "SECTION"
            Height          =   195
            Left            =   870
            TabIndex        =   75
            Top             =   450
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   720
            TabIndex        =   63
            Top             =   840
            Width           =   870
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "MONTH"
            Height          =   195
            Left            =   720
            TabIndex        =   62
            Top             =   1320
            Width           =   600
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2865
         Left            =   -74130
         TabIndex        =   52
         Top             =   630
         Width           =   8415
         Begin MSComCtl2.DTPicker DTPicker_year 
            Height          =   300
            Left            =   3870
            TabIndex        =   53
            Top             =   1140
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   124125187
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "YEAR"
            Height          =   195
            Left            =   3120
            TabIndex        =   54
            Top             =   1170
            Width           =   435
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2835
         Left            =   840
         TabIndex        =   6
         Top             =   660
         Width           =   8415
         Begin VB.Frame Frame6 
            Caption         =   "TYPE"
            Height          =   525
            Left            =   1800
            TabIndex        =   20
            Top             =   1170
            Width           =   4125
            Begin VB.OptionButton opt_dtl_periode 
               Caption         =   "PERIODE"
               Height          =   255
               Left            =   2580
               TabIndex        =   23
               Top             =   210
               Width           =   1185
            End
            Begin VB.OptionButton opt_dtl_monthly 
               Caption         =   "MONTHLY"
               Height          =   255
               Left            =   1200
               TabIndex        =   22
               Top             =   210
               Width           =   1245
            End
            Begin VB.OptionButton opt_dtl_daily 
               Caption         =   "DAILY"
               Height          =   255
               Left            =   180
               TabIndex        =   21
               Top             =   210
               Width           =   885
            End
         End
         Begin VB.TextBox txt_dtl_employee_code 
            Height          =   315
            Left            =   7830
            TabIndex        =   11
            Top             =   810
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Frame fra_dtl_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   7
            Top             =   570
            Visible         =   0   'False
            Width           =   4575
            Begin prj_tpc.vbButton cmd_dtl_browse_employee 
               Height          =   315
               Left            =   1410
               TabIndex        =   49
               Top             =   240
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               BTYPE           =   14
               TX              =   "..."
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
               MICON           =   "frm_rpt_att_detail.frx":86D3
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.TextBox txt_dtl_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   9
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox txt_dtl_nik 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               Height          =   315
               Left            =   0
               MaxLength       =   50
               TabIndex        =   8
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.ComboBox cbo_dtl_employee 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":86EF
            Left            =   1800
            List            =   "frm_rpt_att_detail.frx":86F9
            TabIndex        =   2
            Text            =   "..."
            Top             =   810
            Width           =   1695
         End
         Begin VB.Frame fra_dtl_monthly 
            Height          =   585
            Left            =   1800
            TabIndex        =   14
            Top             =   1650
            Width           =   1875
            Begin MSComCtl2.DTPicker DTPicker_dtl_monthly 
               Height          =   300
               Left            =   90
               TabIndex        =   15
               Top             =   180
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "MM-yyyy"
               Format          =   124125187
               UpDown          =   -1  'True
               CurrentDate     =   39278
            End
         End
         Begin VB.Frame fra_dtl_daily 
            Height          =   585
            Left            =   1800
            TabIndex        =   12
            Top             =   1650
            Width           =   1875
            Begin MSComCtl2.DTPicker DTPicker_dtl_daily 
               Height          =   300
               Left            =   90
               TabIndex        =   13
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   124125187
               CurrentDate     =   39278
            End
         End
         Begin VB.Frame fra_dtl_periode 
            Height          =   585
            Left            =   1800
            TabIndex        =   16
            Top             =   1650
            Width           =   4125
            Begin MSComCtl2.DTPicker DTPicker_dtl_periode_from 
               Height          =   300
               Left            =   150
               TabIndex        =   17
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   124125187
               CurrentDate     =   39278
            End
            Begin MSComCtl2.DTPicker DTPicker_dtl_periode_to 
               Height          =   300
               Left            =   2280
               TabIndex        =   18
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   124125187
               CurrentDate     =   39278
            End
            Begin VB.Label Label6 
               Caption         =   "TO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1950
               TabIndex        =   19
               Top             =   210
               Width           =   285
            End
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   720
            TabIndex        =   10
            Top             =   870
            Width           =   870
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2865
         Left            =   -74160
         TabIndex        =   24
         Top             =   690
         Width           =   8415
         Begin VB.ComboBox cbo_sum_employee 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":870A
            Left            =   1800
            List            =   "frm_rpt_att_detail.frx":8714
            TabIndex        =   35
            Text            =   "..."
            Top             =   780
            Width           =   1695
         End
         Begin VB.TextBox txt_sum_employee_code 
            Height          =   315
            Left            =   7830
            TabIndex        =   31
            Top             =   780
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Frame Frame4 
            Caption         =   "TYPE"
            Height          =   525
            Left            =   1800
            TabIndex        =   25
            Top             =   1140
            Width           =   4125
            Begin VB.OptionButton opt_sum_daily 
               Caption         =   "DAILY"
               Height          =   255
               Left            =   180
               TabIndex        =   28
               Top             =   210
               Width           =   885
            End
            Begin VB.OptionButton opt_sum_monthly 
               Caption         =   "MONTHLY"
               Height          =   255
               Left            =   1200
               TabIndex        =   27
               Top             =   210
               Width           =   1245
            End
            Begin VB.OptionButton opt_sum_periode 
               Caption         =   "PERIODE"
               Height          =   255
               Left            =   2580
               TabIndex        =   26
               Top             =   210
               Width           =   1185
            End
         End
         Begin VB.Frame fra_sum_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   32
            Top             =   540
            Width           =   4575
            Begin VB.TextBox txt_sum_nik 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               Height          =   315
               Left            =   0
               MaxLength       =   50
               TabIndex        =   34
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txt_sum_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   33
               Top             =   240
               Width           =   2415
            End
            Begin prj_tpc.vbButton cmd_sum_browse_employee 
               Height          =   315
               Left            =   1410
               TabIndex        =   50
               Top             =   240
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               BTYPE           =   14
               TX              =   "..."
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
               MICON           =   "frm_rpt_att_detail.frx":8725
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
         End
         Begin VB.Frame fra_sum_monthly 
            Height          =   585
            Left            =   1800
            TabIndex        =   36
            Top             =   1620
            Width           =   1875
            Begin MSComCtl2.DTPicker DTPicker_sum_monthly 
               Height          =   300
               Left            =   90
               TabIndex        =   37
               Top             =   180
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "MM-yyyy"
               Format          =   124125187
               UpDown          =   -1  'True
               CurrentDate     =   39278
            End
         End
         Begin VB.Frame fra_sum_daily 
            Height          =   585
            Left            =   1800
            TabIndex        =   29
            Top             =   1620
            Width           =   1875
            Begin MSComCtl2.DTPicker DTPicker_sum_daily 
               Height          =   300
               Left            =   90
               TabIndex        =   30
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   124125187
               CurrentDate     =   39278
            End
         End
         Begin VB.Frame fra_sum_periode 
            Height          =   585
            Left            =   1800
            TabIndex        =   38
            Top             =   1620
            Width           =   4125
            Begin MSComCtl2.DTPicker DTPicker_sum_periode_from 
               Height          =   300
               Left            =   150
               TabIndex        =   39
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   124125187
               CurrentDate     =   39278
            End
            Begin MSComCtl2.DTPicker DTPicker_sum_periode_to 
               Height          =   300
               Left            =   2280
               TabIndex        =   40
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   124125187
               CurrentDate     =   39278
            End
            Begin VB.Label Label1 
               Caption         =   "TO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1950
               TabIndex        =   41
               Top             =   210
               Width           =   285
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   720
            TabIndex        =   42
            Top             =   840
            Width           =   870
         End
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DEPARTMENT"
      Height          =   195
      Left            =   90
      TabIndex        =   48
      Top             =   1260
      Width           =   1125
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ATTENDANCE REPORT"
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   420
      TabIndex        =   5
      Top             =   870
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_rpt_att_detail.frx":8741
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14850
   End
End
Attribute VB_Name = "frm_rpt_detail_attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCompany As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset
Dim rsDiv As New ADODB.Recordset

Dim int_libur() As Integer
Dim v_access As String
Dim v_dept As String

Private Sub rpt_detail_att()
Dim tgl1 As String, tgl2 As String, SQL As String
Dim str_param_periode As String
Dim a As New frm_rpt

    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
'    v_access = IIf(LOGIN_LEVEL = 100, "''", "(level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND flag_active = 0 order by employee_name")
'    v_dept = IIf(TDBCombo_department.Text = "", "company_code = '" & TDBCombo_company.Text & "'", "company_code = '" & TDBCombo_company.Text & "' AND department_code = '" & TDBCombo_department.Text & "'")
    
    If check_valid_periode Then
        If opt_dtl_daily.Value = True Then
            tgl1 = Format(DTPicker_dtl_daily.Value, "yyyy-MM-dd")
            tgl2 = Format(DTPicker_dtl_daily.Value, "yyyy-MM-dd")
        ElseIf opt_dtl_monthly.Value = True Then
            Call getPeriodeVar(DTPicker_dtl_monthly.Value, tgl1, tgl2)
            tgl1 = Format(tgl1, "yyyy-MM-dd")
            tgl2 = Format(tgl2, "yyyy-MM-dd")
        Else
            tgl1 = Format(DTPicker_dtl_periode_from.Value, "yyyy-MM-dd")
            tgl2 = Format(DTPicker_dtl_periode_to.Value, "yyyy-MM-dd")
        End If
        
        str_param_periode = Format(tgl1, "dd-MMM-yyyy") & " s/d " & Format(tgl2, "dd-MMM-yyyy")
        
        DoEvents
        MousePointer = vbHourglass
        
        If cbo_dtl_employee.ListIndex <> 0 Then
            Call check_late(txt_dtl_employee_code.Text, tgl1, tgl2)
        Else
            SQL = "SELECT employee_code FROM m_employee WHERE flag_active <> 0"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                While Not rs.EOF
                    Call check_late(rs!employee_code, tgl1, tgl2)
                rs.MoveNext
                Wend
            End If
            rs.Close
        End If
        
        MousePointer = vbNormal

        SQL = "call sp_detail_attendance('" & txt_dtl_employee_code.Text & "','" & tgl1 & "','" & tgl2 & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_LEVEL & "','" & TDBCombo_department.Text & "'," & _
                "'" & IIf(LOGIN_LEVEL = 100, LOGIN_FULLNAME, EMPLOYEE_NAME) & "','" & LOGIN_CODE & "')"
        str_file = "\report\rpt_detail_attendance.rpt"
        
        Call a.Show
        a.Caption = "DETAIL ATTENDANCE"
         
        Call a.rpt_view(SQL, str_file, str_param_periode)
    End If

End Sub

Private Sub rpt_detail_sum()
Dim tgl1 As String, tgl2 As String, SQL As String
Dim str_param_periode As String
Dim a As New frm_rpt

    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
'    v_access = IIf(LOGIN_LEVEL = 100, "''", "(level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND flag_active = 0 order by employee_name")
'    v_dept = IIf(TDBCombo_department.Text = "", "company_code = '" & TDBCombo_company.Text & "'", "company_code = '" & TDBCombo_company.Text & "' AND department_code = '" & TDBCombo_department.Text & "'")
    
    If check_valid_periode Then
        If opt_sum_daily.Value = True Then
            tgl1 = Format(DTPicker_sum_daily.Value, "yyyy-MM-dd")
            tgl2 = Format(DTPicker_sum_daily.Value, "yyyy-MM-dd")
        ElseIf opt_sum_monthly.Value = True Then
            Call getPeriodeVar(DTPicker_sum_monthly.Value, tgl1, tgl2)
            tgl1 = Format(tgl1, "yyyy-MM-dd")
            tgl2 = Format(tgl2, "yyyy-MM-dd")
        Else
            tgl1 = Format(DTPicker_sum_periode_from.Value, "yyyy-MM-dd")
            tgl2 = Format(DTPicker_sum_periode_to.Value, "yyyy-MM-dd")
        End If
        
        str_param_periode = Format(tgl1, "dd-MMM-yyyy") & " s/d " & Format(tgl2, "dd-MMM-yyyy")
        
        SQL = "call spr_summary_attendance('" & txt_sum_employee_code.Text & "','" & tgl1 & "','" & tgl2 & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_LEVEL & "','" & TDBCombo_department.Text & "'," & _
                "'" & IIf(LOGIN_LEVEL = 100, LOGIN_FULLNAME, EMPLOYEE_NAME) & "','" & LOGIN_CODE & "')"
        str_file = "\report\rpt_summary_attendance.rpt"
        
        Call a.Show
        a.Caption = "SUMMARY ATTENDANCE"
         
        Call a.rpt_view(SQL, str_file, str_param_periode)
    End If

End Sub

Private Sub rpt_man_hour()
Dim str_param_periode As String
Dim a As New frm_rpt

    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
        
    str_param_periode = "Jan " & Format(DTPicker_Year, "yyyy") & " s/d " & "Des " & Format(DTPicker_Year, "yyyy")
    
    SQL = "call sp_man_hour('" & txt_dtl_employee_code.Text & "','" & Format(DTPicker_Year, "yyyy") & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_LEVEL & "','" & TDBCombo_department.Text & "'," & _
                "'" & IIf(LOGIN_LEVEL = 100, LOGIN_FULLNAME, EMPLOYEE_NAME) & "','" & LOGIN_CODE & "')"
    str_file = "\report\rpt_man_hour.rpt"
    
    Call a.Show
    a.Caption = "SAFE WORKING MAN HOUR RECORD"
     
    Call a.rpt_view(SQL, str_file, str_param_periode)

End Sub

Private Sub rpt_man_hour_emp()
Dim str_param_periode As String
Dim a As New frm_rpt

    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
        
    str_param_periode = Format(DTPicker_periode_from.Value, "dd MMM yyyy") & " s/d " & Format(DTPicker_periode_to.Value, "dd MMM yyyy")
    
    SQL = "call sp_man_hour_emp('" & txt_employee_code.Text & "','" & Format(DTPicker_Year, "yyyy") & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_LEVEL & "','" & TDBCombo_department.Text & "','" & TDBCombo_division.Text & "'," & _
                "'" & IIf(LOGIN_LEVEL = 100, LOGIN_FULLNAME, EMPLOYEE_NAME) & "','" & LOGIN_CODE & "'," & _
                "'" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "')"
    str_file = "\report\rpt_man_hour_emp.rpt"
    
    Call a.Show
    a.Caption = "SAFE WORKING MAN HOUR RECORD"
     
    Call a.rpt_view(SQL, str_file, str_param_periode)

End Sub

Private Sub rpt_sum_att_ot()
Dim str_param_periode As String
Dim dt1 As String, dt2 As String
Dim a As New frm_rpt

    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If
        
    str_param_periode = Format(DTPicker_monthly.Value, "MMMM yyyy")
    Call getPeriode(DTPicker_monthly.Value, DTPicker_periode_from, DTPicker_periode_to)
    
    SQL = "call sp_sum_att_ot('" & txt_employee_code.Text & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_LEVEL & "','" & TDBCombo_department.Text & "','" & TDBCombo_division.Text & "'," & _
                "'" & IIf(LOGIN_LEVEL = 100, LOGIN_FULLNAME, EMPLOYEE_NAME) & "','" & LOGIN_CODE & "'," & _
                "'" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "')"
    str_file = "\report\rpt_summary_calc_att.rpt"
    
    Call a.Show
    a.Caption = "SUMMARY ATTENDANCE & OVERTME CALCULATION FOR  STAF SALARY"
     
    Call a.rpt_view(SQL, str_file, str_param_periode)

End Sub

Private Sub cbo_dtl_employee_Click()
    If cbo_dtl_employee.ListIndex = 0 Then
        fra_dtl_employee.Visible = False
    Else
        fra_dtl_employee.Visible = True
    End If
    
    txt_dtl_employee_code = "": txt_dtl_nik = "": txt_dtl_employee_name = ""
End Sub

Private Sub cbo_sum_employee_Click()
    If cbo_sum_employee.ListIndex = 0 Then
        fra_sum_employee.Visible = False
    Else
        fra_sum_employee.Visible = True
    End If
    
    txt_sum_employee_code = "": txt_sum_nik = "": txt_sum_employee_name = ""
End Sub

Private Sub cbo_employee_Click()
    If cbo_employee.ListIndex = 0 Then
        fra_employee.Visible = False
    Else
        fra_employee.Visible = True
    End If
    
    txt_employee_code = "": txt_nik = "": txt_employee_name = ""
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Function check_valid_periode() As Boolean
check_valid_periode = True

    If SSTab1.Tab = 0 Then
        If cbo_dtl_employee.ListIndex = 1 And Trim(txt_dtl_nik.Text) = "" Then
            MsgBox "Employee is not selected!", vbOKOnly + vbInformation, headerMSG
            cmd_dtl_browse_employee.SetFocus
            check_valid_periode = False
            Exit Function
        End If
    Else
        If cbo_sum_employee.ListIndex = 1 And Trim(txt_sum_nik.Text) = "" Then
            MsgBox "Employee is not selected!", vbOKOnly + vbInformation, headerMSG
            cmd_sum_browse_employee.SetFocus
            check_valid_periode = False
            Exit Function
        End If
    End If
End Function

Private Sub cmdPrint_Click()
Dim rsemp As New ADODB.Recordset
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        Call rpt_detail_att
    ElseIf SSTab1.Tab = 1 Then
        Call rpt_detail_sum
    ElseIf SSTab1.Tab = 2 Then
        If opt_monthly.Value Then
            Call getPeriode(DTPicker_monthly.Value, DTPicker_periode_from, DTPicker_periode_to)
                        
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            ProgressBar1.Min = 0
            
            If rsemp.State Then rsemp.Close
            SQL = "SELECT employee_code FROM m_employee " & _
                  "WHERE CASE WHEN '" & txt_employee_code.Text & "' = '' THEN " & _
                           "CASE WHEN '" & TDBCombo_division.Text & "' <> '' THEN company_code = '" & TDBCombo_company.Text & "' AND department_code = '" & TDBCombo_department.Text & "' AND division_code = '" & TDBCombo_division.Text & "' " & _
                           "ELSE CASE WHEN '" & TDBCombo_department.Text & "' <> '' THEN company_code = '" & TDBCombo_company.Text & "' AND department_code = '" & TDBCombo_department.Text & "' " & _
                           "ELSE company_code = '" & TDBCombo_company.Text & "' END END " & _
                         "ELSE employee_code = '" & txt_employee_code.Text & "' END"
            rsemp.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rsemp.RecordCount > 0 Then
                ProgressBar1.Max = rsemp.RecordCount
                
                rsemp.MoveFirst
                While Not rsemp.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
                    SQL = "DELETE FROM t_spl_auto " & _
                          "WHERE employee_code = '" & rsemp.Fields(0).Value & "' " & _
                            "AND (DATE(date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " & _
                                  "AND '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "')"
                    CnG.Execute SQL
                    
                    Call frm_trans_salary_process.auto_overtime(rsemp.Fields(0).Value, DTPicker_periode_from.Value, DTPicker_periode_to.Value)
                    Call check_late(rsemp.Fields(0).Value, DTPicker_periode_from.Value, DTPicker_periode_to.Value)
                    rsemp.MoveNext
                Wend
            End If
            rsemp.Close
            DoEvents
            ProgressBar1.Visible = False
            
            Call rpt_sum_att_ot
        ElseIf opt_periode.Value Then
            Call rpt_man_hour_emp
        End If
    ElseIf SSTab1.Tab = 3 Then
        Call rpt_man_hour
    End If
    
End Sub

Private Sub Form_Load()
    Call load_data_company
    Call load_data_user_access(Me)
    
    SSTab1.Tab = 0
    
    opt_dtl_daily.Value = True
    
    cbo_dtl_employee.ListIndex = 0
    cbo_sum_employee.ListIndex = 0
    cbo_employee.ListIndex = 0
    
    ProgressBar1.Visible = False
    
    Call createGridKar
    timer1.Enabled = True
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub opt_dtl_daily_Click()
    fra_dtl_daily.Visible = True
    fra_dtl_monthly.Visible = False
    fra_dtl_periode.Visible = False
    
    DTPicker_dtl_daily.Value = Now
End Sub

Private Sub opt_dtl_monthly_Click()
    fra_dtl_daily.Visible = False
    fra_dtl_monthly.Visible = True
    fra_dtl_periode.Visible = False
    
    DTPicker_dtl_monthly.Value = Now
End Sub

Private Sub opt_dtl_periode_Click()
    fra_dtl_daily.Visible = False
    fra_dtl_monthly.Visible = False
    fra_dtl_periode.Visible = True
    
    DTPicker_dtl_periode_from.Value = Now
    DTPicker_dtl_periode_to.Value = Now
End Sub

Private Sub opt_sum_daily_Click()
    fra_sum_daily.Visible = True
    fra_sum_monthly.Visible = False
    fra_sum_periode.Visible = False
    
    DTPicker_sum_daily.Value = Now
End Sub

Private Sub opt_sum_monthly_Click()
    fra_sum_daily.Visible = False
    fra_sum_monthly.Visible = True
    fra_sum_periode.Visible = False
    
    DTPicker_sum_monthly.Value = Now
End Sub

Private Sub opt_sum_periode_Click()
    fra_sum_daily.Visible = False
    fra_sum_monthly.Visible = False
    fra_sum_periode.Visible = True
    
    DTPicker_sum_periode_from.Value = Now
    DTPicker_sum_periode_to.Value = Now
End Sub

Private Sub opt_monthly_Click()
    fra_monthly.Visible = True
    fra_periode.Visible = False
    
    DTPicker_monthly.Value = Now
End Sub

Private Sub opt_periode_Click()
    fra_monthly.Visible = False
    fra_periode.Visible = True
    
    DTPicker_periode_from.Value = Now
    DTPicker_periode_to.Value = Now
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        opt_dtl_daily.Value = True
    ElseIf SSTab1.Tab = 1 Then
        opt_sum_daily.Value = True
    ElseIf SSTab1.Tab = 2 Then
        opt_monthly.Value = True
    ElseIf SSTab1.Tab = 3 Then
        DTPicker_Year.Value = Now
    End If
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If
    
    Call load_data_department
End Sub

Private Sub timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub load_data_department()
    If rsDept.State Then rsDept.Close
    SQL = "select * from m_department where company_code = '" & TDBCombo_company.Text & "' order by department_code"
    rsDept.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_department.RowSource = rsDept
End Sub

Private Sub load_data_division()
    If rsDiv.State Then rsDiv.Close
    SQL = "select * from m_division where department_code = '" & TDBCombo_department.Text & "' order by division_code"
    rsDiv.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_division.RowSource = rsDiv
End Sub

Private Sub TDBCombo_department_Change()
    If TDBCombo_department.Text = "" Then txt_department_name.Text = ""
End Sub

Private Sub TDBCombo_department_ItemChange()
    If TDBCombo_department.ApproxCount > 0 Then
        TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
        txt_department_name = TDBCombo_department.Columns("department_name").Value
        
        Call load_data_division
    End If
End Sub

Private Sub TDBCombo_division_Change()
    If TDBCombo_division.Text = "" Then txt_division_name.Text = ""
End Sub

Private Sub TDBCombo_division_ItemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
        txt_division_name = TDBCombo_division.Columns("division_name").Value
    End If
End Sub

Private Sub createGridKar()
   With LynxGrid2
      .AddColumn "Employee Code", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "employee_code", 2000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear
        
        vParam = IIf(DEPARTMENT_CODE <> "" And DIVISION_CODE = "", "a.department_code = '" & DEPARTMENT_CODE & "'", IIf(DEPARTMENT_CODE = "" And DIVISION_CODE = "", "a.company_code = '" & COMPANY_CODE & "'", "a.department_code = '" & DEPARTMENT_CODE & "' AND a.division_code = '" & DIVISION_CODE & "'"))
        
        If SSTab1.Tab = 0 Then
            If LOGIN_LEVEL = 100 Then
                SQL = "select nik,employee_name,employee_code " & _
                         "from m_employee a " & _
                         "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                            "AND (nik LIKE '%" & txt_dtl_nik.Text & "%' " & _
                                "OR employee_name LIKE '%" & txt_dtl_nik.Text & "%')"
            Else
                SQL = "select nik,employee_name,employee_code " & _
                         "from m_employee a " & _
                         "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                            "AND " & vParam & " " & _
                            "AND (nik LIKE '%" & txt_dtl_nik.Text & "%' " & _
                                "OR employee_name LIKE '%" & txt_dtl_nik.Text & "%') " & _
                            "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
            End If
        ElseIf SSTab1.Tab = 1 Then
            If LOGIN_LEVEL = 100 Then
                SQL = "select nik,employee_name,employee_code " & _
                         "from m_employee a " & _
                         "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                            "AND (nik LIKE '%" & txt_sum_nik.Text & "%' " & _
                                "OR employee_name LIKE '%" & txt_sum_nik.Text & "%')"
            Else
                SQL = "select nik,employee_name,employee_code " & _
                         "from m_employee a " & _
                         "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                            "AND " & vParam & " " & _
                            "AND (nik LIKE '%" & txt_sum_nik.Text & "%' " & _
                                "OR employee_name LIKE '%" & txt_sum_nik.Text & "%') " & _
                            "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
            End If
        ElseIf SSTab1.Tab = 2 Then
            If LOGIN_LEVEL = 100 Then
                SQL = "select nik,employee_name,employee_code " & _
                         "from m_employee a " & _
                         "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                            "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                                "OR employee_name LIKE '%" & txt_nik.Text & "%')"
            Else
                SQL = "select nik,employee_name,employee_code " & _
                         "from m_employee a " & _
                         "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                            "AND " & vParam & " " & _
                            "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                                "OR employee_name LIKE '%" & txt_nik.Text & "%') " & _
                            "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
            End If
        End If
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid2.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME & vbTab & rs!employee_code
                rs.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                
                If SSTab1.Tab = 0 Then
                    txt_dtl_employee_code.Text = rs!employee_code
                    txt_dtl_employee_name.Text = rs!EMPLOYEE_NAME
                    txt_dtl_nik.Text = rs!nik
                ElseIf SSTab1.Tab = 1 Then
                    txt_sum_employee_code.Text = rs!employee_code
                    txt_sum_employee_name.Text = rs!EMPLOYEE_NAME
                    txt_sum_nik.Text = rs!nik
                ElseIf SSTab1.Tab = 2 Then
                    txt_employee_code.Text = rs!employee_code
                    txt_employee_name.Text = rs!EMPLOYEE_NAME
                    txt_nik.Text = rs!nik
                End If
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
            If SSTab1.Tab = 0 Then
                txt_dtl_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
                txt_dtl_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
                txt_dtl_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
            ElseIf SSTab1.Tab = 1 Then
                txt_sum_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
                txt_sum_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
                txt_sum_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
            ElseIf SSTab1.Tab = 2 Then
                txt_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
                txt_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
                txt_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
            End If
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

Private Sub txt_dtl_nik_Change()
    If txt_dtl_nik.Text = "" Then
        txt_dtl_employee_code.Text = ""
        txt_dtl_employee_name.Text = ""
    End If
End Sub

Private Sub txt_sum_nik_Change()
    If txt_sum_nik.Text = "" Then
        txt_sum_employee_code.Text = ""
        txt_sum_employee_name.Text = ""
    End If
End Sub

Private Sub txt_nik_Change()
    If txt_nik.Text = "" Then
        txt_employee_code.Text = ""
        txt_employee_name.Text = ""
    End If
End Sub

Private Sub txt_dtl_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub txt_sum_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub txt_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmd_dtl_browse_employee_Click()
    isiGridKar (1)
End Sub

Private Sub cmd_sum_browse_employee_Click()
    isiGridKar (1)
End Sub

Private Sub cmd_browse_employee_Click()
    isiGridKar (1)
End Sub

