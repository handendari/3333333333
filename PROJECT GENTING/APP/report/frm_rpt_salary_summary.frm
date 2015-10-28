VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_rpt_summary_salary 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SUMMARY SALARY REPORT"
   ClientHeight    =   7245
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   10560
   Icon            =   "frm_rpt_salary_summary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Text            =   "frm_rpt_salary_summary.frx":000C
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3030
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   690
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4755
      Left            =   270
      TabIndex        =   3
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8387
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "MONTHLY"
      TabPicture(0)   =   "frm_rpt_salary_summary.frx":0012
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra_monthly"
      Tab(0).Control(1)=   "fra_process_monthly"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "PAYROLL CALCULATION"
      TabPicture(1)   =   "frm_rpt_salary_summary.frx":002E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_periode"
      Tab(1).Control(1)=   "fra_process_periode"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "SALARY CALCULATION SUMMARY"
      TabPicture(2)   =   "frm_rpt_salary_summary.frx":004A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "REQUEST FOR PAYMENT"
      TabPicture(3)   =   "frm_rpt_salary_summary.frx":0066
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   4005
         Left            =   -74220
         TabIndex        =   104
         Top             =   450
         Width           =   8625
         Begin VB.TextBox txt_auth22 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4380
            TabIndex        =   131
            Top             =   2940
            Width           =   1275
         End
         Begin VB.TextBox txt_auth12 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4380
            TabIndex        =   128
            Top             =   2250
            Width           =   1275
         End
         Begin VB.TextBox txt_check2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   510
            TabIndex        =   125
            Top             =   2940
            Width           =   1275
         End
         Begin VB.TextBox txt_prepare2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   510
            TabIndex        =   120
            Top             =   2250
            Width           =   1275
         End
         Begin VB.TextBox txtSisaGaji 
            Height          =   285
            Left            =   6420
            TabIndex        =   124
            Text            =   "Text2"
            Top             =   270
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.TextBox txt_payment_for2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   118
            Top             =   1860
            Width           =   6255
         End
         Begin VB.TextBox txt_payment_for1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   117
            Top             =   1530
            Width           =   6255
         End
         Begin VB.TextBox txt_payment_to 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   116
            Top             =   1140
            Width           =   6255
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":0082
            Left            =   3600
            List            =   "frm_rpt_salary_summary.frx":008C
            Locked          =   -1  'True
            TabIndex        =   105
            Text            =   "..."
            Top             =   750
            Width           =   1335
         End
         Begin VB.TextBox txt_hr_executive2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   121
            Top             =   2250
            Width           =   2385
         End
         Begin VB.TextBox txt_hr_executive2_title 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   123
            Top             =   2550
            Width           =   2385
         End
         Begin VB.TextBox txt_hr_manager2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   126
            Top             =   2940
            Width           =   2385
         End
         Begin VB.TextBox txt_hr_manager2_title 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   127
            Top             =   3240
            Width           =   2385
         End
         Begin VB.TextBox txt_finance2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   129
            Top             =   2250
            Width           =   2385
         End
         Begin VB.TextBox txt_finance2_title 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   130
            Top             =   2550
            Width           =   2385
         End
         Begin VB.TextBox txt_presdir2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   132
            Top             =   2940
            Width           =   2385
         End
         Begin VB.TextBox txt_presdir2_title 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   133
            Top             =   3240
            Width           =   2385
         End
         Begin MSComCtl2.DTPicker DTPicker_month_rfp_from 
            Height          =   300
            Left            =   1800
            TabIndex        =   106
            Top             =   750
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   225247235
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_month_rfp_to 
            Height          =   300
            Left            =   5040
            TabIndex        =   107
            Top             =   750
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   225247235
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_month_rfp 
            Height          =   300
            Left            =   1800
            TabIndex        =   108
            Top             =   390
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   225247235
            CurrentDate     =   39278
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Payment For"
            Height          =   195
            Left            =   540
            TabIndex        =   122
            Top             =   1560
            Width           =   885
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Payment To"
            Height          =   195
            Left            =   540
            TabIndex        =   119
            Top             =   1170
            Width           =   855
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Bulan"
            Height          =   195
            Left            =   540
            TabIndex        =   115
            Top             =   390
            Width           =   405
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Periode"
            Height          =   195
            Left            =   540
            TabIndex        =   114
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label45 
            Caption         =   "* yyyy-MM"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   3540
            TabIndex        =   113
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   540
            TabIndex        =   112
            Top             =   2580
            Width           =   570
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   540
            TabIndex        =   111
            Top             =   3300
            Width           =   570
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   4410
            TabIndex        =   110
            Top             =   2580
            Width           =   570
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   4410
            TabIndex        =   109
            Top             =   3270
            Width           =   570
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3915
         Left            =   750
         TabIndex        =   51
         Top             =   510
         Width           =   8625
         Begin VB.TextBox txt_auth2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4380
            TabIndex        =   84
            Top             =   2940
            Width           =   1275
         End
         Begin VB.TextBox txt_auth1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4380
            TabIndex        =   80
            Top             =   2250
            Width           =   1275
         End
         Begin VB.TextBox txt_check 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   510
            TabIndex        =   76
            Top             =   2940
            Width           =   1275
         End
         Begin VB.TextBox txt_prepare 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   510
            TabIndex        =   64
            Top             =   2250
            Width           =   1275
         End
         Begin VB.TextBox txt_title_presdir 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   86
            Top             =   3240
            Width           =   2385
         End
         Begin VB.TextBox txt_presdir 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   85
            Top             =   2940
            Width           =   2385
         End
         Begin VB.TextBox txt_title_finance 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   82
            Top             =   2550
            Width           =   2385
         End
         Begin VB.TextBox txt_finance 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   81
            Top             =   2250
            Width           =   2385
         End
         Begin VB.TextBox txt_title_hr_manager 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   78
            Top             =   3240
            Width           =   2385
         End
         Begin VB.TextBox txt_hr_manager 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   77
            Top             =   2940
            Width           =   2385
         End
         Begin VB.TextBox txt_title_hr_executive 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   74
            Top             =   2550
            Width           =   2385
         End
         Begin VB.TextBox txt_hr_executive 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   65
            Top             =   2250
            Width           =   2385
         End
         Begin VB.TextBox txt_address 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   63
            Top             =   1830
            Width           =   2385
         End
         Begin VB.CommandButton Command2 
            Caption         =   "DAY COUNT"
            Height          =   495
            Left            =   7020
            TabIndex        =   53
            Top             =   -90
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":009A
            Left            =   3600
            List            =   "frm_rpt_salary_summary.frx":00A4
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "..."
            Top             =   1380
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   300
            Left            =   1800
            TabIndex        =   54
            Top             =   1380
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   111476739
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   300
            Left            =   5040
            TabIndex        =   55
            Top             =   1380
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   111476739
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   300
            Left            =   1800
            TabIndex        =   62
            Top             =   930
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   111476739
            CurrentDate     =   39278
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_Level 
            Height          =   375
            Left            =   1800
            OleObjectBlob   =   "frm_rpt_salary_summary.frx":00B2
            TabIndex        =   60
            Top             =   450
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txt_level 
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
            Height          =   285
            Left            =   3540
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   59
            Top             =   450
            Visible         =   0   'False
            Width           =   2415
         End
         Begin MSAdodcLib.Adodc Adodc_level 
            Height          =   375
            Left            =   2220
            Top             =   420
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   4410
            TabIndex        =   87
            Top             =   3270
            Width           =   570
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   4410
            TabIndex        =   83
            Top             =   2580
            Width           =   570
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   540
            TabIndex        =   79
            Top             =   3300
            Width           =   570
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   540
            TabIndex        =   75
            Top             =   2580
            Width           =   570
         End
         Begin VB.Label Label21 
            Caption         =   "* yyyy-MM"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   3540
            TabIndex        =   70
            Top             =   930
            Width           =   1125
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Addressed To"
            Height          =   195
            Left            =   540
            TabIndex        =   66
            Top             =   1860
            Width           =   990
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Level"
            Height          =   195
            Left            =   540
            TabIndex        =   58
            Top             =   510
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Periode"
            Height          =   195
            Left            =   540
            TabIndex        =   57
            Top             =   1380
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Bulan"
            Height          =   195
            Left            =   540
            TabIndex        =   56
            Top             =   930
            Width           =   405
         End
      End
      Begin VB.Frame fra_monthly 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   7
         Top             =   960
         Width           =   8415
         Begin VB.ComboBox cbo_monthly_company 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":2012
            Left            =   1800
            List            =   "frm_rpt_salary_summary.frx":201C
            TabIndex        =   13
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.Frame fra_monthly_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   9
            Top             =   960
            Width           =   4575
            Begin VB.CommandButton cmd_monthly_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1440
               TabIndex        =   12
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txt_monthly_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   11
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox txt_monthly_employee_code 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   0
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   10
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.ComboBox cbo_monthly_employee 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":202F
            Left            =   1800
            List            =   "frm_rpt_salary_summary.frx":2039
            TabIndex        =   8
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_monthly 
            Height          =   300
            Left            =   1800
            TabIndex        =   14
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   110755843
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Perusahaan"
            Height          =   195
            Left            =   720
            TabIndex        =   17
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Bulan"
            Height          =   195
            Left            =   720
            TabIndex        =   16
            Top             =   1620
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Karyawan"
            Height          =   195
            Left            =   720
            TabIndex        =   15
            Top             =   1200
            Width           =   705
         End
      End
      Begin VB.Frame fra_process_monthly 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   37
         Top             =   960
         Width           =   8415
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Processing, Please Wait..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2640
            TabIndex        =   38
            Top             =   1080
            Width           =   2730
         End
      End
      Begin VB.Frame fra_periode 
         Height          =   4215
         Left            =   -74160
         TabIndex        =   18
         Top             =   360
         Width           =   8415
         Begin VB.TextBox txt_auth21 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4380
            TabIndex        =   100
            Top             =   2550
            Width           =   1275
         End
         Begin VB.TextBox txt_auth11 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4380
            TabIndex        =   96
            Top             =   1860
            Width           =   1275
         End
         Begin VB.TextBox txt_check1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   510
            TabIndex        =   92
            Top             =   2550
            Width           =   1275
         End
         Begin VB.TextBox txt_prepare1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   510
            TabIndex        =   88
            Top             =   1860
            Width           =   1275
         End
         Begin VB.TextBox txt_hr_executive1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   89
            Top             =   1860
            Width           =   2385
         End
         Begin VB.TextBox txt_hr_executive1_title 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   91
            Top             =   2160
            Width           =   2385
         End
         Begin VB.TextBox txt_hr_manager1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   93
            Top             =   2550
            Width           =   2385
         End
         Begin VB.TextBox txt_hr_manager1_title 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   95
            Top             =   2850
            Width           =   2385
         End
         Begin VB.TextBox txt_finance1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   97
            Top             =   1860
            Width           =   2385
         End
         Begin VB.TextBox txt_finance1_title 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   99
            Top             =   2160
            Width           =   2385
         End
         Begin VB.TextBox txt_presdir1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   101
            Top             =   2550
            Width           =   2385
         End
         Begin VB.TextBox txt_presdir1_title 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5670
            TabIndex        =   103
            Top             =   2850
            Width           =   2385
         End
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   255
            Left            =   270
            TabIndex        =   71
            Top             =   3480
            Visible         =   0   'False
            Width           =   7875
            _ExtentX        =   13891
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.TextBox txt_periode_employee_code 
            Height          =   285
            Left            =   5130
            TabIndex        =   67
            Text            =   "Text2"
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cbo_periode_department 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":204A
            Left            =   1800
            List            =   "frm_rpt_salary_summary.frx":2054
            TabIndex        =   49
            Text            =   "..."
            Top             =   300
            Width           =   1695
         End
         Begin VB.ComboBox cbo_periode_to 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":2067
            Left            =   3600
            List            =   "frm_rpt_salary_summary.frx":2071
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "..."
            Top             =   1470
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "DAY COUNT"
            Height          =   495
            Left            =   0
            TabIndex        =   24
            Top             =   -240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cbo_periode_employee 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":207F
            Left            =   1800
            List            =   "frm_rpt_salary_summary.frx":2089
            TabIndex        =   23
            Text            =   "..."
            Top             =   720
            Width           =   1695
         End
         Begin VB.Frame fra_periode_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   375
            Left            =   3600
            TabIndex        =   19
            Top             =   690
            Width           =   4575
            Begin VB.TextBox txt_periode_nik 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   0
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   22
               Top             =   0
               Width           =   1335
            End
            Begin VB.TextBox txt_periode_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   21
               Top             =   0
               Width           =   2415
            End
            Begin VB.CommandButton cmd_periode_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1470
               TabIndex        =   20
               Top             =   0
               Width           =   375
            End
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_from 
            Height          =   300
            Left            =   1800
            TabIndex        =   27
            Top             =   1470
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   111935491
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_to 
            Height          =   300
            Left            =   5040
            TabIndex        =   28
            Top             =   1470
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   111935491
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   1800
            TabIndex        =   43
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   111935491
            CurrentDate     =   39278
         End
         Begin VB.Frame fra_periode_department 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   3600
            TabIndex        =   46
            Top             =   210
            Width           =   4695
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
               Height          =   285
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               MultiLine       =   -1  'True
               TabIndex        =   47
               Top             =   90
               Width           =   2415
            End
            Begin TrueOleDBList60.TDBCombo TDBCombo_department 
               Height          =   375
               Left            =   0
               OleObjectBlob   =   "frm_rpt_salary_summary.frx":209A
               TabIndex        =   48
               Top             =   90
               Width           =   1815
            End
            Begin MSAdodcLib.Adodc Adodc_department 
               Height          =   375
               Left            =   690
               Top             =   90
               Visible         =   0   'False
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
         End
         Begin VB.ComboBox cbo_periode_company 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":4003
            Left            =   6270
            List            =   "frm_rpt_salary_summary.frx":400D
            TabIndex        =   25
            Text            =   "..."
            Top             =   840
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   540
            TabIndex        =   102
            Top             =   2190
            Width           =   570
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   540
            TabIndex        =   98
            Top             =   2910
            Width           =   570
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   4410
            TabIndex        =   94
            Top             =   2190
            Width           =   570
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Jabatan"
            Height          =   195
            Left            =   4410
            TabIndex        =   90
            Top             =   2880
            Width           =   570
         End
         Begin VB.Label Label23 
            Caption         =   "Sending Mail, Please Wait...."
            Height          =   225
            Left            =   330
            TabIndex        =   73
            Top             =   3780
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Label14"
            Height          =   255
            Left            =   6060
            TabIndex        =   72
            Top             =   3780
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label20 
            Caption         =   "* yyyy-MM"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   3600
            TabIndex        =   69
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Month"
            Height          =   195
            Left            =   540
            TabIndex        =   44
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Periode"
            Height          =   195
            Left            =   540
            TabIndex        =   31
            Top             =   1470
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Department"
            Height          =   195
            Left            =   540
            TabIndex        =   30
            Top             =   330
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Employee"
            Height          =   195
            Left            =   540
            TabIndex        =   29
            Top             =   690
            Width           =   690
         End
      End
      Begin VB.Frame fra_process_periode 
         Height          =   2475
         Left            =   -74160
         TabIndex        =   35
         Top             =   960
         Width           =   8415
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   1770
            Width           =   8205
            _ExtentX        =   14473
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            Height          =   345
            Left            =   450
            TabIndex        =   41
            Top             =   1230
            Width           =   2385
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            Height          =   405
            Left            =   3660
            TabIndex        =   40
            Top             =   1230
            Width           =   4575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Processing, Please Wait..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2670
            TabIndex        =   36
            Top             =   270
            Width           =   2730
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   270
      TabIndex        =   0
      Top             =   5880
      Width           =   10095
      Begin VB.CommandButton cmd_Print 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5970
         Picture         =   "frm_rpt_salary_summary.frx":4020
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint_Rekap 
         Caption         =   "&Recapitu lation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4920
         Picture         =   "frm_rpt_salary_summary.frx":45AA
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdSlipKetapang 
         Caption         =   "&Slip Ketapang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1290
         Picture         =   "frm_rpt_salary_summary.frx":4B34
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmd_print_slip 
         Caption         =   "&Slip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2760
         Picture         =   "frm_rpt_salary_summary.frx":50BE
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Calculation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3840
         Picture         =   "frm_rpt_salary_summary.frx":5648
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_print_sum_list 
         Caption         =   "&Sum. List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5520
         Picture         =   "frm_rpt_salary_summary.frx":5BD2
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmd_send_mail 
         Caption         =   "Send &Mail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   7800
         Picture         =   "frm_rpt_salary_summary.frx":615C
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   840
         Top             =   360
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   8820
         Picture         =   "frm_rpt_salary_summary.frx":66E6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1230
      OleObjectBlob   =   "frm_rpt_salary_summary.frx":6C70
      TabIndex        =   5
      Top             =   690
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1080
      Top             =   690
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9000
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "LAPORAN SALARY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4065
      TabIndex        =   68
      Top             =   0
      Width           =   2955
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Perusahaan"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   690
      Width           =   855
   End
End
Attribute VB_Name = "frm_rpt_summary_salary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' *****************************************************************************
' Required declaration of the vbSendMail component (withevents is optional)
' You also need a reference to the vbSendMail component in the Project References
' *****************************************************************************
'Private WithEvents poSendMail As clsSendMail
'
'' misc local vars
'Dim bAuthLogin      As Boolean
'Dim bPopLogin       As Boolean
'Dim bHtml           As Boolean
'Dim MyEncodeType    As ENCODE_METHOD
'Dim etPriority      As MAIL_PRIORITY
'Dim bReceipt        As Boolean
'Dim keluar          As Boolean

Dim int_libur() As Integer
'Dim pub_employee_code, pub_employee_name, pub_company_code, pub_company_name, pub_email As String
'Public int_sent_false, int_sent_true As Integer
'
'Public str_smtp, str_sender_mail, str_sender_name As String
'Public str_smtp_port, str_username, str_sender_pwd As String

Dim rs_employee As New ADODB.Recordset
Dim rs_proses As New ADODB.Recordset
Dim v_proses As Integer

Private Sub cbo_monthly_company_Click()
If cbo_monthly_company.ListIndex = 0 Then
    cbo_monthly_employee.ListIndex = 0
    cbo_monthly_employee.Enabled = False
ElseIf cbo_monthly_company.ListIndex = 1 Then
    cbo_monthly_employee.ListIndex = 0
    cbo_monthly_employee.Enabled = True
End If
End Sub

Private Sub cbo_monthly_employee_Click()
If cbo_monthly_employee.ListIndex = 0 Then
    fra_monthly_employee.Visible = False
Else
    fra_monthly_employee.Visible = True
    txt_monthly_employee_code = "": txt_monthly_employee_name = ""
End If
End Sub

Private Sub cbo_periode_department_Click()
TDBCombo_department.Text = ""
txt_department_name.Text = ""
    
If cbo_periode_department.ListIndex = 0 Then
    fra_periode_department.Visible = False
Else
    fra_periode_department.Visible = True
End If
End Sub

Private Sub cbo_periode_employee_Click()
If cbo_periode_employee.ListIndex = 0 Then
    fra_periode_employee.Visible = False
Else
    fra_periode_employee.Visible = True
    txt_periode_employee_code = "": txt_periode_employee_name = ""
End If
End Sub

Private Sub cmd_daily_browse_employee_Click()
frm_lookup_mst_employee.public_int_mode = 6
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_monthly_browse_employee_Click()
frm_lookup_mst_employee.public_int_mode = 78
frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_periode_browse_employee_Click()
frm_lookup_mst_employee.public_int_mode = 79
frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_Print_Click()

If SSTab1.Tab = 2 Then
    Call DTPicker4_Change
    frm_lookup_level.public_int = 2
    frm_lookup_level.Show
Else
    Call DTPicker_month_rfp_Change
    frm_lookup_level.public_int = 3
    frm_lookup_level.Show
'    Call print_rfp
End If

End Sub

Private Sub cmd_print_slip_Click()
Dim rs_proses As New ADODB.Recordset
Dim v_proses As Integer
Dim str_sql As String

If check_validate_tdbcombo(TDBCombo_company) = False Then
    MsgBox "No Company selected!", vbInformation, headerMSG
    Exit Sub
End If

Call DTPicker1_Change

'If SSTab1.Tab = 1 Then
    If check_valid_periode Then
'        Call rpt_periode(1)
        Call DTPicker1_Change
        frm_lookup_level.public_int = 4
        frm_lookup_level.Show
    End If
'ElseIf SSTab1.Tab = 0 Then
'    If check_valid_monthly Then
'        Call rpt_monthly(1)
'    End If

'End If
End Sub

Private Sub cmd_send_mail_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If

    Call DTPicker1_Change
    
    If check_valid_periode Then
        Call DTPicker1_Change
        frm_lookup_level.public_int = 7
        frm_lookup_level.Show
    End If

End Sub

'Private Sub send_mail_monthly(ByVal j As Integer)
'Dim str_sql, str_param_periode, str_file, str1, str_file_out As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim dt1, dt2 As Date
'Dim d1, d2 As String
'Dim a As New frm_rpt
'
'If j = 0 Then
'    str_file = "\report\rpt_53_1.rpt"
'ElseIf j = 1 Then
'    str_file = "\report\rpt_53.rpt"
'End If
'
'If cbo_monthly_company.ListIndex = 0 Then
'    int_flag_company = 0
'    str_company_code = "-"
'ElseIf cbo_monthly_company.ListIndex = 1 Then
'    int_flag_company = 1
'    str_company_code = TDBCombo_company.Columns("company_code").Value
'End If
'
'
'If int_flag_company = 0 Then
'    int_flag_employee = 0
'    str_employee_code = "-"
'    str1 = "select * from m_employee where ifnull(flag_active,0)=1 order by employee_code asc"
'ElseIf int_flag_company = 1 Then
'    If cbo_monthly_employee.ListIndex = 0 Then
'        int_flag_employee = 0
'        str_employee_code = "-"
'        str1 = "select * from m_employee where company_code='" & str_company_code & "' and ifnull(flag_active,0)=1 order by employee_code asc"
'    ElseIf cbo_monthly_employee.ListIndex = 1 Then
'        int_flag_employee = 1
'        str_employee_code = txt_monthly_employee_code
'        str1 = "select * from m_employee where employee_code='" & str_employee_code & "' and ifnull(flag_active,0)=1 order by employee_code asc"
'    End If
'End If
'
'If rs_employee.State = 1 Then rs_employee.Close
'rs_employee.Open str1, CnG, adOpenStatic, adLockReadOnly
'If rs_employee.RecordCount > 0 Then rs_employee.MoveFirst
'
'dt1 = Format(DTPicker_monthly.Value, "yyyy-MM-01")
'dt2 = DateAdd("m", 1, dt1): dt2 = DateAdd("d", -1, dt2)
'd1 = Format(dt1, "yyyy-mm-dd")
'd2 = Format(dt2, "yyyy-mm-dd")
'
'While Not rs_employee.EOF
'
'    str_sql = "call spr_attendance_54('" & d1 & "','" & d2 & "'," _
'            & 1 & ",'" & rs_employee!COMPANY_CODE & "'," & 1 & ",'" & rs_employee!employee_code & "')"
'    str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
'
'    '-- creating pdf
'    str_file_out = App.Path & "\mail\slip_" & Format(DTPicker_monthly, "yyyymm") & "_" & rs_employee!employee_code & ".pdf"
'    Call rpt_auto_pdf(str_sql, str_file, str_file_out, str_param_periode)
'
'    pub_employee_code = rs_employee!employee_code
'    pub_employee_name = rs_employee!EMPLOYEE_NAME
'    pub_company_code = rs_employee!COMPANY_CODE
'    pub_company_name = rs_employee!company_name
'    pub_email = "" & rs_employee!email
'
'    Call send_mail("Salary slip " & str_param_periode, "Salary slip " & str_param_periode & vbCrLf & vbCrLf _
'                    & "TTD" & vbCrLf & str_sender_name, str_file_out)
'    '--
'
'    rs_employee.MoveNext
'Wend
'
'Call show_msg
'End Sub

'Private Sub send_mail_periode(ByVal j As Integer)
'Dim str_sql, str_param_periode, str_file, str1, str_file_out As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim a As New frm_rpt
'
'If j = 0 Then
'    str_file = "\report\rpt_53_1.rpt"
'ElseIf j = 1 Then
'    str_file = "\report\rpt_53.rpt"
'End If
'
'If cbo_periode_company.ListIndex = 0 Then
'    int_flag_company = 0
'    str_company_code = "-"
'ElseIf cbo_periode_company.ListIndex = 1 Then
'    int_flag_company = 1
'    str_company_code = TDBCombo_company.Columns("company_code").Value
'End If
'
'If int_flag_company = 0 Then
'    int_flag_employee = 0
'    str_employee_code = "-"
'    str1 = "select * from m_employee where ifnull(flag_active,0)=1 order by employee_code asc"
'
'ElseIf int_flag_company = 1 Then
'    If cbo_periode_employee.ListIndex = 0 Then
'        int_flag_employee = 0
'        str_employee_code = "-"
'        str1 = "select * from m_employee where company_code='" & str_company_code & "' and ifnull(flag_active,0)=1 order by employee_code asc"
'    ElseIf cbo_periode_employee.ListIndex = 1 Then
'        int_flag_employee = 1
'        str_employee_code = txt_periode_employee_code
'        str1 = "select * from m_employee where employee_code='" & str_employee_code & "' and ifnull(flag_active,0)=1 order by employee_code asc"
'    End If
'End If
'
'If rs_employee.State = 1 Then rs_employee.Close
'rs_employee.Open str1, CnG, adOpenStatic, adLockReadOnly
'If rs_employee.RecordCount > 0 Then rs_employee.MoveFirst
'
'While Not rs_employee.EOF
'
'    If cbo_periode_to.ListIndex = 0 Then
'        str_sql = "call spr_attendance_54('" _
'            & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" _
'            & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "'," _
'            & 1 & ",'" & rs_employee!COMPANY_CODE & "'," & 1 & ",'" & rs_employee!employee_code & "')"
'        str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
'    ElseIf cbo_periode_to.ListIndex = 1 Then
'        str_sql = "call spr_attendance_54('" _
'            & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" _
'            & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "'," _
'            & 1 & ",'" & rs_employee!COMPANY_CODE & "'," & 1 & ",'" & rs_employee!employee_code & "')"
'        str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
'            & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
'    End If
'
'
'    '-- creating pdf
'    str_file_out = App.Path & "\mail\slip_" & Format(DTPicker_periode_from, "yyyymm") & "_" & rs_employee!employee_code & ".pdf"
'    Call rpt_auto_pdf(str_sql, str_file, str_file_out, str_param_periode)
'
'    pub_employee_code = rs_employee!employee_code
'    pub_employee_name = rs_employee!EMPLOYEE_NAME
'    pub_company_code = rs_employee!COMPANY_CODE
'    pub_company_name = rs_employee!company_name
'    pub_email = "" & rs_employee!email
'
'    Call send_mail("Salary slip " & str_param_periode, "Salary slip " & str_param_periode & vbCrLf & vbCrLf _
'                    & "TTD" & vbCrLf & str_sender_name, str_file_out)
'    '--
'
'    rs_employee.MoveNext
'Wend
'
'Call show_msg
'End Sub

'Private Sub show_msg()
'MsgBox "There are " & int_sent_true & " mail are sent successfully!" & vbCrLf _
'    & int_sent_false & " are fails!" & vbCrLf _
'    & "For more detail info, let see the 'log mail'!", vbInformation, headerMSG
'End Sub

Private Sub rpt_auto_pdf(ByVal sql_proc As String, ByVal rpt_file As String, _
ByVal str_file_out As String, ByVal str_param As String)
Dim CrApp As New CRAXDRT.Application
Dim CrRep As New CRAXDRT.Report
Dim AdoRs As New ADODB.Recordset

    AdoRs.Open sql_proc, CnG, adOpenDynamic, adLockBatchOptimistic
    Set CrRep = CrApp.OpenReport(App.Path & rpt_file)
    CrRep.DiscardSavedData
    CrRep.Database.Tables(1).SetDataSource AdoRs, 3
    CrRep.ParameterFields.GetItemByName("p_periode").AddCurrentValue str_param
    
    '---
    CrRep.ExportOptions.DestinationType = crEDTDiskFile
    CrRep.ExportOptions.FormatType = crEFTPortableDocFormat
    CrRep.ExportOptions.DiskFileName = str_file_out
    CrRep.Export False
End Sub

Private Sub cmdPrint_Rekap_Click()

Call DTPicker1_Change

frm_lookup_level.public_int = 1
frm_lookup_level.Show
End Sub

Private Sub cmdSlipKetapang_Click()
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer
Dim strsql As String
Dim rsemployee As New ADODB.Recordset

int_process = vbNo

str_file = "\report\rpt_slip_new.rpt" 'rpt_hardys_salary2.rpt"

d1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
d2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")

If cbo_periode_company.ListIndex = 0 Then
    int_flag_company = 0
    str_company_code = "-"
ElseIf cbo_periode_company.ListIndex = 1 Then
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
End If

If int_flag_company = 0 Then
    int_flag_employee = 0
    str_employee_code = "-"
ElseIf int_flag_company = 1 Then
    If cbo_periode_employee.ListIndex = 0 Then
        int_flag_employee = 0
        str_employee_code = "-"
    ElseIf cbo_periode_employee.ListIndex = 1 Then
        int_flag_employee = 1
        str_employee_code = txt_periode_employee_code
    End If
End If

'--
    If cbo_periode_to.ListIndex = 0 Then
        str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
            & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
            & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "')"
        'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
    ElseIf cbo_periode_to.ListIndex = 1 Then
        str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
            & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
            & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "')"
        'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
            & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
    End If
'--


Text1 = str_sql

Call a.Show

a.Caption = "SALARY SLIP REPORT"
Call a.rpt_print(str_sql, str_file, str_param_periode)

fra_process_periode.Visible = False
fra_periode.Visible = True
End Sub

Private Sub Command1_Click()
'MsgBox TDBDate1.MinDate & " - " & TDBDate1.MaxDate _
& vbCr & TDBDate1.Month _
& vbCr & TDBDate1.Weekday


'MsgBox DTPicker1.DayOfWeek
'ReDim int_libur(5)
Dim a, b As Date

'a = DTPicker_periode_from.Value
'b = DateAdd("m", 1, a)

a = Format(DTPicker_periode_from.Value, "yyyy-MM-01")
b = DateAdd("m", 1, a)
'DTPicker_periode_to.Value = DateAdd("m", 1, DTPicker_periode_from.Value)

'MsgBox DTPicker_periode_from.Value & vbCr _
& DateDiff("d", DTPicker_periode_from.Value, DTPicker_periode_to.Value)

MsgBox a & vbCr _
& DateDiff("d", a, b)
End Sub

'Private Sub cmdSlipKetapang_Click()
''MsgBox UBound(int_libur)
'
''Call get_holiday("2007-07")
'
'Dim rs As New ADODB.Recordset
'Dim cmd As New ADODB.Command
'
'cmd.ActiveConnection = CnG
'cmd.CommandText = Text1.Text
'rs.CursorLocation = adUseClient
'rs.Open cmd, , adOpenStatic, adLockReadOnly
'
'MsgBox rs.RecordCount
'End Sub


Private Sub get_holiday(ByVal i As String)
'DTPicker1.Value = i & "-01"
End Sub

Private Sub periode_date_event()
If cbo_periode_to.ListIndex = 0 Then
    DTPicker_periode_to.Visible = False
    DTPicker_periode_to.Value = DTPicker_periode_from.Value
Else
    DTPicker_periode_to.Visible = True
    DTPicker_periode_to.Value = DTPicker_periode_from.Value
End If
End Sub

Private Sub cbo_periode_to_Click()
Call periode_date_event
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Function check_valid_periode() As Boolean
check_valid_periode = True

'validate employee
If cbo_periode_employee.ListIndex = 1 And Trim(txt_periode_employee_code) = "" Then
    MsgBox "Employee is not selected!", vbOKOnly + vbInformation, headerMSG
    cmd_periode_browse_employee.SetFocus
    check_valid_periode = False
    Exit Function
End If
End Function

Private Function check_valid_monthly() As Boolean
check_valid_monthly = True

'validate employee
If cbo_monthly_employee.ListIndex = 1 And Trim(txt_monthly_employee_code) = "" Then
    MsgBox "Employee is not selected!", vbOKOnly + vbInformation, headerMSG
    cmd_monthly_browse_employee.SetFocus
    check_valid_monthly = False
    Exit Function
End If
End Function

Private Function check_valid_daily() As Boolean
check_valid_daily = True

'validate employee
'If cbo_daily_employee.ListIndex = 1 And Trim(txt_daily_employee_code) = "" Then
'    MsgBox "Employee is not selected!", vbOKOnly + vbInformation, headerMSG
'    cmd_daily_browse_employee.SetFocus
'    check_valid_daily = False
'    Exit Function
'End If
End Function

'Private Sub rpt_periode(ByVal j As Integer)
'
''Dim str_sql, str_param_periode, str_file, str1, str2  As String
''Dim int_flag_company As Integer, str_company_code As String
''Dim int_flag_employee As Integer, str_employee_code As String
''Dim a As New frm_rpt
''Dim d1, d2, dx As Date
''Dim int_process As Integer
''Dim strsql As String
''Dim rsemployee As New ADODB.Recordset
''
'''+++++++++++++++++++++++++++++++++ Check Temp Salary Proses ++++++++++++++++++++++++++++++++++++++
''str_sql = "SELECT salary_proses FROM temp_sal_proses WHERE company_code = '" & TDBCombo_company.Text & "'"
''rs_proses.Open str_sql, CnG, adOpenForwardOnly, adLockReadOnly
''    v_proses = rs_proses!salary_proses
''rs_proses.Close
''
''    If v_proses = 0 Then
''        MsgBox "There are some data that has been changed." & Chr(13) & _
''            "Please re salary process!", vbExclamation, headerMSG
''        Exit Sub
''    End If
'''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''
''int_process = vbNo
''
''If j = 0 Then
''    str_file = "\report\rpt_salary_full.rpt" 'rpt_hardys_salary1.rpt"
''ElseIf j = 2 Then
''    str_file = "\report\rpt_salary_full.rpt" 'rpt_hardys_salary1.rpt"
''ElseIf j = 1 Then
''    str_file = "\report\rpt_slip_new.rpt" 'rpt_hardys_salary2.rpt"
''End If
''d1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
''d2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
''
''If cbo_periode_company.ListIndex = 0 Then
''    int_flag_company = 0
''    str_company_code = "-"
''ElseIf cbo_periode_company.ListIndex = 1 Then
''    int_flag_company = 1
''    str_company_code = TDBCombo_company.Columns("company_code").Value
''End If
''
''If int_flag_company = 0 Then
''    int_flag_employee = 0
''    str_employee_code = "-"
''ElseIf int_flag_company = 1 Then
''    If cbo_periode_employee.ListIndex = 0 Then
''        int_flag_employee = 0
''        str_employee_code = "-"
''    ElseIf cbo_periode_employee.ListIndex = 1 Then
''        int_flag_employee = 1
''        str_employee_code = txt_periode_employee_code
''    End If
''End If
''
''
''    If j = 0 Or j = 2 Then
''    '--
''        If cbo_periode_to.ListIndex = 0 Then
''            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
''                & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
''                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
''                & "'" & txt_hr_executive1.Text & "','" & txt_hr_executive1_title.Text & "'," _
''                & "'" & txt_hr_manager1.Text & "','" & txt_hr_manager1_title.Text & "'," _
''                & "'" & txt_finance1.Text & "','" & txt_finance1_title.Text & "'," _
''                & "'" & txt_presdir1.Text & "','" & txt_presdir1_title.Text & "')"
''
''            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
''        ElseIf cbo_periode_to.ListIndex = 1 Then
''            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
''                & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
''                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
''                & "'" & txt_hr_executive1.Text & "','" & txt_hr_executive1_title.Text & "'," _
''                & "'" & txt_hr_manager1.Text & "','" & txt_hr_manager1_title.Text & "'," _
''                & "'" & txt_finance1.Text & "','" & txt_finance1_title.Text & "'," _
''                & "'" & txt_presdir1.Text & "','" & txt_presdir1_title.Text & "')"
''
''            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
''                & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
''        End If
''    '--
''    ElseIf j = 1 Then
''    '--
''        If cbo_periode_to.ListIndex = 0 Then
''            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
''                & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
''                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
''                & "'" & txt_hr_executive1.Text & "','" & txt_hr_executive1_title.Text & "'," _
''                & "'" & txt_hr_manager1.Text & "','" & txt_hr_manager1_title.Text & "'," _
''                & "'" & txt_finance1.Text & "','" & txt_finance1_title.Text & "'," _
''                & "'" & txt_presdir1.Text & "','" & txt_presdir1_title.Text & "')"
''            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
''        ElseIf cbo_periode_to.ListIndex = 1 Then
''            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
''                & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
''                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
''                & "'" & txt_hr_executive1.Text & "','" & txt_hr_executive1_title.Text & "'," _
''                & "'" & txt_hr_manager1.Text & "','" & txt_hr_manager1_title.Text & "'," _
''                & "'" & txt_finance1.Text & "','" & txt_finance1_title.Text & "'," _
''                & "'" & txt_presdir1.Text & "','" & txt_presdir1_title.Text & "')"
''
''            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
''                & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
''        End If
''    '--
''    End If
''
''str_param_periode = Format(DTPicker1.Value, "mmmm yyyy")
''
''
''Text1 = str_sql
''
''Call a.Show
''a.Caption = "SALARY REPORT"
''Call a.rpt_view(str_sql, str_file, str_param_periode)
''
''fra_process_periode.Visible = False
''fra_periode.Visible = True
'End Sub

Private Sub rpt_monthly(ByVal j As Integer)
Dim str_sql, str_param_periode, str_file, str1, str2 As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer
Dim strsql As String
Dim rsemployee As New ADODB.Recordset

int_process = vbNo

If j = 0 Then
    str_file = "\report\rpt_53_1_ebiz.rpt"
ElseIf j = 2 Then
    str_file = "\report\rpt_53_1_ebiz_list.rpt"
ElseIf j = 1 Then
    'str_file = "\report\rpt_53_ebiz.rpt"
    str_file = "\report\spt_slip_new.rpt"
End If

If cbo_monthly_company.ListIndex = 0 Then
    int_flag_company = 0
    str_company_code = "-"
ElseIf cbo_monthly_company.ListIndex = 1 Then
    int_flag_company = 1
    str_company_code = "" & TDBCombo_company.Columns("company_code").Value
End If

If int_flag_company = 0 Then
    int_flag_employee = 0
    str_employee_code = "-"
ElseIf int_flag_company = 1 Then
    If cbo_monthly_employee.ListIndex = 0 Then
        int_flag_employee = 0
        str_employee_code = "-"
    ElseIf cbo_monthly_employee.ListIndex = 1 Then
        int_flag_employee = 1
        str_employee_code = txt_monthly_employee_code
    End If
End If

If j = 0 Or j = 2 Then
    '--
    str_sql = "call spr_salary_ebiz_sum('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "','" & LOGIN_CODE & "')"
    'str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
    '--
ElseIf j = 1 Then
    '--
    str_sql = "call spr_salary_ebiz_sli('" & Format(DTPicker_monthly.Value, "yyyy-MM") & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "','" & LOGIN_CODE & "')"
        
    'str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
    '--
End If

Text1 = str_sql

Call a.Show
a.Caption = "SALARY REPORT"
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub cmdPrint_Click()
If check_validate_tdbcombo(TDBCombo_company) = False Then
    MsgBox "No Company selected!", vbInformation, headerMSG
    Exit Sub
End If

'Call DTPicker1_Change

If SSTab1.Tab = 1 Then
    If check_valid_periode Then
'        Call rpt_periode(0)
        Call DTPicker1_Change
        frm_lookup_level.public_int = 5
        frm_lookup_level.Show
    End If
ElseIf SSTab1.Tab = 0 Then
    If check_valid_monthly Then
        Call rpt_monthly(0)
    End If

End If
End Sub

'Private Sub cmd_print_sum_list_Click()
'If check_validate_tdbcombo(TDBCombo_company) = False Then
'    MsgBox "No Company selected!", vbInformation, headerMSG
'    Exit Sub
'End If
'
'If SSTab1.Tab = 1 Then
'    If check_valid_periode Then
'        Call rpt_periode(2)
'    End If
'ElseIf SSTab1.Tab = 0 Then
'    If check_valid_monthly Then
'        Call rpt_monthly(2)
'    End If
'
'End If
'End Sub

Private Sub Command3_Click()
Adodc1.ConnectionString = strConn
Adodc1.RecordSource = Text1.Text
Adodc1.Refresh

MsgBox Adodc1.Recordset.RecordCount
End Sub
Private Sub DTPicker_month_rfp_Change()
Dim strsql As String
Dim rsbulan As New ADODB.Recordset

    strsql = "select periode_from,periode_to from h_d_salary " & _
            "WHERE company_code = '" & TDBCombo_company.Text & "' AND left(`month`,7) = '" & Format(DTPicker_month_rfp.Value, "yyyy-MM") & "'"
    rsbulan.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsbulan.RecordCount = 0 Then
        MsgBox "Transaksi Bulan Ini Belum di Proses..!!", vbInformation, "Informasi"
        DTPicker4.SetFocus
        Exit Sub
    End If
    
    DTPicker_month_rfp_from.Value = rsbulan!periode_from
    DTPicker_month_rfp_to.Value = rsbulan!periode_to
    Combo2.ListIndex = 1
    rsbulan.Close
    
'    strsql = "SELECT SUM(CASE WHEN salary_code = 'SU-37' THEN salary_value ELSE 0 END) AS sisa_gaji " _
'            & "FROM m_employee a JOIN m_company c ON a.company_code = c.company_code " _
'            & "LEFT JOIN h_salary b ON a.employee_code=b.employee_code AND LEFT(b.month,7) = '" & Format(DTPicker_month_rfp.Value, "yyyy-MM") & "' " _
'            & "JOIN (SELECT employee_code FROM h_attendance WHERE DATE(att_date) BETWEEN '" & Format(DTPicker_month_rfp_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_month_rfp_to.Value, "yyyy-MM-dd") & "' " _
'                    & "GROUP BY employee_code) d ON a.employee_code = d.employee_code " _
'            & "WHERE a.company_code = '" & TDBCombo_company.Text & "' AND " _
'                    & "(level_code = ANY (SELECT access_level_code FROM t_user_access_level " _
'                        & "WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND a.flag_active = 1 " _
'            & "GROUP BY a.company_code"
'    rsbulan.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'    If rsbulan.RecordCount > 0 Then
'        txtSisaGaji.Text = rsbulan!sisa_gaji
'    End If
'    rsbulan.Close
End Sub

Private Sub DTPicker_month_rfp_Validate(Cancel As Boolean)
Dim strsql As String
Dim rsbulan As New ADODB.Recordset

    strsql = "select periode_from,periode_to from h_d_salary " & _
            "WHERE company_code = '" & TDBCombo_company.Text & "' AND left(`month`,7) = '" & Format(DTPicker_month_rfp.Value, "yyyy-MM") & "'"
    rsbulan.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsbulan.RecordCount = 0 Then
        MsgBox "Transaksi Bulan Ini Belum di Proses..!!", vbInformation, "Informasi"
        DTPicker4.SetFocus
        Exit Sub
    End If
    
    DTPicker_month_rfp_from.Value = rsbulan!periode_from
    DTPicker_month_rfp_to.Value = rsbulan!periode_to
    Combo2.ListIndex = 1
    rsbulan.Close
    
'    strsql = "SELECT SUM(CASE WHEN salary_code = 'SU-37' THEN ROUND(salary_value) ELSE 0 END) AS sisa_gaji " _
'            & "FROM m_employee a JOIN m_company c ON a.company_code = c.company_code " _
'            & "LEFT JOIN h_salary b ON a.employee_code=b.employee_code AND LEFT(b.month,7) = '" & Format(DTPicker_month_rfp.Value, "yyyy-MM") & "' " _
'            & "JOIN (SELECT employee_code FROM h_attendance WHERE DATE(att_date) BETWEEN '" & Format(DTPicker_month_rfp_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_month_rfp_to.Value, "yyyy-MM-dd") & "' " _
'                    & "GROUP BY employee_code) d ON a.employee_code = d.employee_code " _
'            & "WHERE a.company_code = '" & TDBCombo_company.Text & "' AND " _
'                    & "(level_code = ANY (SELECT access_level_code FROM t_user_access_level " _
'                        & "WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND a.flag_active = 1 " _
'            & "GROUP BY a.company_code"
'    rsbulan.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'    If rsbulan.RecordCount > 0 Then
'        txtSisaGaji.Text = rsbulan!sisa_gaji
'    End If
'    rsbulan.Close
End Sub

Private Sub DTPicker4_Change()
Dim strsql As String
Dim rsbulan As New ADODB.Recordset

    strsql = "select periode_from,periode_to from h_d_salary " & _
            "WHERE company_code = '" & TDBCombo_company.Text & "' AND left(`month`,7) = '" & Format(DTPicker4.Value, "yyyy-MM") & "'"
    rsbulan.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsbulan.RecordCount = 0 Then
        MsgBox "Transaksi Bulan Ini Belum di Proses..!!", vbInformation, "Informasi"
        DTPicker4.SetFocus
        Exit Sub
    End If
    
    DTPicker2.Value = rsbulan!periode_from
    DTPicker3.Value = rsbulan!periode_to
    Combo2.ListIndex = 1
End Sub

Private Sub DTPicker1_Change()
Dim strsql As String
Dim rsbulan As New ADODB.Recordset

    strsql = "select periode_from,periode_to from h_d_salary " & _
            "WHERE company_code = '" & TDBCombo_company.Text & "' AND left(`month`,7) = '" & Format(DTPicker1.Value, "yyyy-MM") & "'"
    rsbulan.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsbulan.RecordCount = 0 Then
        MsgBox "Transaksi Bulan Ini Belum di Proses..!!", vbInformation, "Informasi"
        DTPicker1.SetFocus
        Exit Sub
    End If
    
    DTPicker_periode_from.Value = rsbulan!periode_from
    DTPicker_periode_to.Value = rsbulan!periode_to
    cbo_periode_to.ListIndex = 1
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
Dim strsql As String
Dim rsbulan As New ADODB.Recordset

    strsql = "select periode_from,periode_to from h_d_salary " & _
            "WHERE company_code = '" & TDBCombo_company.Text & "' AND left(`month`,7) = '" & Format(DTPicker1.Value, "yyyy-MM") & "'"
    rsbulan.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsbulan.RecordCount = 0 Then
        MsgBox "Transaksi Bulan Ini Belum di Proses..!!", vbInformation, "Informasi"
        DTPicker1.SetFocus
        Exit Sub
    End If
    
    DTPicker_periode_from.Value = rsbulan!periode_from
    DTPicker_periode_to.Value = rsbulan!periode_to
    cbo_periode_to.ListIndex = 1
End Sub

Private Sub DTPicker4_Validate(Cancel As Boolean)
Dim strsql As String
Dim rsbulan As New ADODB.Recordset

    strsql = "select periode_from,periode_to from h_d_salary " & _
            "WHERE company_code = '" & TDBCombo_company.Text & "' AND left(`month`,7) = '" & Format(DTPicker4.Value, "yyyy-MM") & "'"
    rsbulan.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsbulan.RecordCount = 0 Then
        MsgBox "Transaksi Bulan Ini Belum di Proses..!!", vbInformation, "Informasi"
        DTPicker4.SetFocus
        Exit Sub
    End If
    
    DTPicker2.Value = rsbulan!periode_from
    DTPicker3.Value = rsbulan!periode_to
    Combo2.ListIndex = 1
End Sub

Private Sub Form_Load()
Adodc_company.ConnectionString = strConn
Adodc_department.ConnectionString = strConn
Adodc_level.ConnectionString = strConn

Label10.Caption = ""
Label11.Caption = ""
Label10.Visible = False
Label11.Visible = False
ProgressBar1.Visible = False
fra_process_periode.Visible = False
DTPicker1.Value = Date

Call load_data_company
'Call load_data_setting_mail

DTPicker_periode_from.Value = Now
DTPicker_periode_to.Value = Now
DTPicker_monthly.Value = Now
DTPicker2.Value = Now
DTPicker3.Value = Now
DTPicker4.Value = Now
DTPicker_month_rfp.Value = Now
DTPicker_month_rfp_from.Value = Now
DTPicker_month_rfp_to.Value = Now

cbo_periode_to.ListIndex = 0
cbo_periode_company.ListIndex = 1
cbo_periode_employee.ListIndex = 0

cbo_monthly_company.ListIndex = 1
cbo_monthly_employee.ListIndex = 0

timer1.Enabled = True
SSTab1.TabVisible(0) = False

cbo_periode_to.ListIndex = 1
cbo_periode_department.ListIndex = 0

cmd_Print.Visible = False
SSTab1.Tab = 1
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub load_data_monthly_company()
'Adodc_monthly_company.RecordSource = "select * from m_company order by company_code"
'Adodc_monthly_company.Refresh
'
'TDBCombo_monthly_company.RowSource = Adodc_monthly_company
End Sub

Private Sub TDBCombo_karyawan_ItemChange()
'If Not (TDBCombo_karyawan.ApproxCount > 0 And TDBCombo_karyawan.Bookmark > 0) Then Exit Sub
'
'TDBCombo_karyawan.Text = TDBCombo_karyawan.Columns("kode_karyawan").Value
'txt_nama_karyawan = TDBCombo_karyawan.Columns("nama_karyawan").Value
End Sub

Private Sub set_company_option()
'If opt_per_company Then
'    TDBGrid1.Enabled = True
'ElseIf opt_all Then
'    TDBGrid1.Enabled = False
'End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
    cmd_print_slip.Visible = True
    cmdPrint.Visible = True
    cmdPrint_Rekap.Visible = True
    cmd_Print.Visible = False
    cmd_send_mail.Visible = True
Else
    cmd_print_slip.Visible = False
    cmdPrint.Visible = False
    cmdPrint_Rekap.Visible = False
    cmd_Print.Visible = True
    cmd_send_mail.Visible = False
End If
End Sub

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
End If

Call load_data_department
Call load_data_level
End Sub

Private Sub TDBCombo_monthly_company_ItemChange()
'If TDBCombo_monthly_company.ApproxCount > 0 Then
'    TDBCombo_monthly_company.Text = TDBCombo_monthly_company.Columns("company_code").Value
'    txt_monthly_company_name = TDBCombo_monthly_company.Columns("company_name").Value
'End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
End Sub

Private Sub Timer1_Timer()
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
If LOGIN_LEVEL = 100 Then
'    cbo_daily_company.Enabled = True
    cbo_monthly_company.Enabled = True
    cbo_periode_company.Enabled = True
Else
'    cbo_daily_company.Enabled = False
    cbo_monthly_company.Enabled = False
    cbo_periode_company.Enabled = False
End If
End Sub

'Private Sub load_data_setting_mail()
'Dim rs1 As New ADODB.Recordset
'Dim strsql As String
'
'strsql = "select * from s_mail where s_number = 1"
'rs1.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'If rs1.RecordCount > 0 Then
'    str_sender_mail = rs1!s_sender_email
'    str_sender_name = rs1!s_sender_name
'    str_smtp = rs1!smtp
'    str_smtp_port = rs1!PORT
'    str_username = rs1!Username
'    str_sender_pwd = RC4DeCryptASC(rs1!Password, pEncryptionPassword)
'Else
'    str_sender_mail = ""
'    str_sender_name = ""
'    str_smtp = ""
'    str_smtp_port = ""
'    str_username = ""
'    str_sender_pwd = ""
'End If
'End Sub

'Private Sub send_mail(ByVal str_subject As String, ByVal str_msg As String, ByVal str_attc As String)
'Set poSendMail = New clsSendMail
'
'With poSendMail
'
'    ' **************************************************************************
'    ' Optional properties for sending email, but these should be set first
'    ' if you are going to use them
'    ' **************************************************************************
'
'    .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
'    .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
'    .Delimiter = ";"                            ' Optional, default = ";" (semicolon)
'
'    ' **************************************************************************
'    ' Basic properties for sending email
'    ' **************************************************************************
'    .SMTPHost = str_smtp
'    .SMTPPort = str_smtp_port
'    '.SMTPHost = "mail.solusisentraldata.com"
'    .from = str_sender_mail
'    .FromDisplayName = str_sender_name
'    .Recipient = rs_employee.Fields("email").Value
'    .RecipientDisplayName = rs_employee.Fields("employee_nick_name").Value
''        .CcRecipient = "CcRecipient"
''        .CcDisplayName = "CcDisplayName"
''        .BccRecipient = "BccRecipient"
'    '.ReplyToAddress = txtFrom.Text              ' Optional, used when different than 'From' address
'    .Subject = str_subject                  ' Optional
'    .Message = str_msg                      ' Optional
'    .Attachment = str_attc          ' Optional, separate multiple entries with delimiter character
'
'    ' **************************************************************************
'    ' Additional Optional properties, use as required by your application / environment
'    ' **************************************************************************
''        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
''        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
''        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
''        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
''        .Receipt = bReceipt                         ' Optional, default = FALSE
''        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
''        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
''        .Username = "send.ktc@solusisentraldata.com"  ' Optional, default = Null String
''        .Password = "sepatukuda2011"                            ' Optional, default = Null String, value is NOT saved
'        .Username = str_username                    ' Optional, default = Null String
'        .Password = str_sender_pwd
''        .POP3Host = txtPopServer
''        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
'
'    ' **************************************************************************
'    ' Advanced Properties, change only if you have a good reason to do so.
'    ' **************************************************************************
'    ' .ConnectTimeout = 10                      ' Optional, default = 10
'    ' .ConnectRetry = 5                         ' Optional, default = 5
'    ' .MessageTimeout = 60                      ' Optional, default = 60
'    ' .PersistentSettings = True                ' Optional, default = TRUE
'    ' .SMTPPort = 25                            ' Optional, default = 25
'
'    ' **************************************************************************
'    ' OK, all of the properties are set, send the email...
'    ' **************************************************************************
'    ' .Connect                                  ' Optional, use when sending bulk mail
'    .Send                                       ' Required
'    ' .Disconnect                               ' Optional, use when sending bulk mail
''        txtServer.Text = .SMTPHost                  ' Optional, re-populate the Host in case
'                                                ' MX look up was used to find a host    End With
'End With
'End Sub
'
'Private Sub poSendMail_SendFailed(Explanation As String)
'Dim rs1 As New ADODB.Recordset
'
'rs1.Open "select * from h_send_mail where employee_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic
'
'CnG.BeginTrans
'
'With rs1
'    .AddNew
'    .Fields("date").Value = Now
'    .Fields("employee_code").Value = pub_employee_code
'    .Fields("employee_name").Value = pub_employee_name
'    .Fields("email").Value = pub_email
'    .Fields("sent_status").Value = 0
'    .Fields("description").Value = "Your attempt to send mail failed for the following reason(s): " _
'                                    & vbCrLf & Explanation
'    .Update
'End With
'
'CnG.CommitTrans
'
'int_sent_false = int_sent_false + 1
'End Sub
'
'Private Sub poSendMail_SendSuccesful()
'Dim rs1 As New ADODB.Recordset
'
'rs1.Open "select * from h_send_mail where employee_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic
'
'CnG.BeginTrans
'
'With rs1
'    .AddNew
'    .Fields("date").Value = Now
'    .Fields("employee_code").Value = pub_employee_code
'    .Fields("employee_name").Value = pub_employee_name
'    .Fields("email").Value = pub_email
'    .Fields("sent_status").Value = 1
'    .Fields("description").Value = "sent successfully"
'    .Update
'End With
'
'CnG.CommitTrans
'
'int_sent_true = int_sent_true + 1
'End Sub
'
'Private Sub poSendMail_Status(Status As String)
''    vbSendMail 'Status Event'
''    lstStatus.AddItem Status
''    lstStatus.ListIndex = lstStatus.ListCount - 1
''    lstStatus.ListIndex = -1
'End Sub
'
'Private Sub poSendMail_Progress(lPercentCompete As Long)
''   vbSendMail 'Progress Event'
''   lblProgress = lPercentCompete & "% complete"
'End Sub

Private Sub set_data_department(ByVal str_code As String)
On Error Resume Next

Adodc_department.Recordset.MoveFirst
Adodc_department.Recordset.Find ("department_code='" & str_code & "'")   ', 0, adSearchForward, 1)
If Not (Adodc_department.Recordset.EOF = True Or Adodc_department.Recordset.BOF = True) Then
    TDBCombo_department.Bookmark = Adodc_department.Recordset.AbsolutePosition
    Call TDBCombo_department_ItemChange
Else
    TDBCombo_department.Text = ""
End If
End Sub

Private Sub load_data_department()
TDBCombo_department.Text = "": txt_department_name = ""

Adodc_department.RecordSource = "select * from m_department where company_code='" _
& TDBCombo_company.Columns("company_code").Value & "' order by department_code"
Adodc_department.Refresh

TDBCombo_department.RowSource = Adodc_department
End Sub

Private Sub TDBCombo_department_ItemChange()
If TDBCombo_department.ApproxCount > 0 Then
    TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
    txt_department_name = TDBCombo_department.Columns("department_name").Value
End If
End Sub

Private Sub txt_kode_bank_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_nama_bank_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub load_data_level()
TDBCombo_level.Text = "": txt_level = ""

Adodc_level.RecordSource = "select a.access_level_code, b.name " _
& "from t_user_access_level a JOIN m_akses_level_group b on a.access_level_code = b.code " _
& "where a.allow_access <> 0 AND a.level_code = '" & LOGIN_CODE & "'"
Adodc_level.Refresh

TDBCombo_level.RowSource = Adodc_level
End Sub

Private Sub TDBCombo_level_ItemChange()
If TDBCombo_level.ApproxCount > 0 Then
    TDBCombo_level.Text = TDBCombo_level.Columns("access_level_code").Value
    txt_level = TDBCombo_level.Columns("name").Value
End If
End Sub
