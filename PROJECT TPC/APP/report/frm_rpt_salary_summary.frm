VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_rpt_summary_salary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SALARY REPORT"
   ClientHeight    =   7545
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   10560
   Icon            =   "frm_rpt_salary_summary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   3465
      Left            =   4650
      TabIndex        =   47
      Top             =   3420
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
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   7380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   41
      Text            =   "frm_rpt_salary_summary.frx":058A
      Top             =   660
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   270
      TabIndex        =   34
      Top             =   5910
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   840
         Top             =   360
      End
      Begin prj_tpc.vbButton cmdExit 
         Height          =   705
         Left            =   8970
         TabIndex        =   35
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
         MICON           =   "frm_rpt_salary_summary.frx":0590
         PICN            =   "frm_rpt_salary_summary.frx":05AC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmd_print_slip 
         Height          =   705
         Left            =   3270
         TabIndex        =   36
         Top             =   270
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Slip"
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
         MICON           =   "frm_rpt_salary_summary.frx":163E
         PICN            =   "frm_rpt_salary_summary.frx":165A
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
         Left            =   4260
         TabIndex        =   37
         Top             =   270
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Summary"
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
         MICON           =   "frm_rpt_salary_summary.frx":26EC
         PICN            =   "frm_rpt_salary_summary.frx":2708
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmd_send_mail 
         Height          =   705
         Left            =   7950
         TabIndex        =   38
         Top             =   300
         Visible         =   0   'False
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
         MICON           =   "frm_rpt_salary_summary.frx":379A
         PICN            =   "frm_rpt_salary_summary.frx":37B6
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
      Left            =   3030
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      Top             =   990
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   270
      TabIndex        =   0
      Top             =   1380
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "MONTHLY"
      TabPicture(0)   =   "frm_rpt_salary_summary.frx":4848
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra_process_monthly"
      Tab(0).Control(1)=   "fra_monthly"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "PERIODE"
      TabPicture(1)   =   "frm_rpt_salary_summary.frx":4864
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fra_process_periode"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra_periode"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fra_periode 
         Height          =   2595
         Left            =   780
         TabIndex        =   17
         Top             =   930
         Width           =   8415
         Begin VB.ComboBox cbo_periode_company 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":4880
            Left            =   7800
            List            =   "frm_rpt_salary_summary.frx":488A
            TabIndex        =   40
            Text            =   "..."
            Top             =   180
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox cbo_periode_employee 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":489D
            Left            =   1800
            List            =   "frm_rpt_salary_summary.frx":48A7
            TabIndex        =   20
            Text            =   "..."
            Top             =   780
            Width           =   1695
         End
         Begin VB.ComboBox cbo_periode_to 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":48B8
            Left            =   3600
            List            =   "frm_rpt_salary_summary.frx":48C2
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "..."
            Top             =   1560
            Width           =   1335
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
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   390
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_from 
            Height          =   300
            Left            =   1800
            TabIndex        =   24
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92078083
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_to 
            Height          =   300
            Left            =   5040
            TabIndex        =   25
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92078083
            CurrentDate     =   39278
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_department 
            Height          =   375
            Left            =   1800
            OleObjectBlob   =   "frm_rpt_salary_summary.frx":48D0
            TabIndex        =   26
            Top             =   420
            Width           =   1695
         End
         Begin MSAdodcLib.Adodc Adodc_department 
            Height          =   375
            Left            =   2490
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
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   1800
            TabIndex        =   27
            Top             =   1170
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "MM-yyyy"
            Format          =   92078083
            CurrentDate     =   39278
         End
         Begin VB.Frame fra_periode_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   21
            Top             =   540
            Width           =   4575
            Begin prj_tpc.vbButton cmd_periode_browse_employee 
               Height          =   315
               Left            =   1410
               TabIndex        =   46
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
               MICON           =   "frm_rpt_salary_summary.frx":6839
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.TextBox txt_periode_employee_code 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   4350
               TabIndex        =   39
               Top             =   240
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.TextBox txt_periode_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   23
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox txt_periode_nik 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               Height          =   315
               Left            =   0
               MaxLength       =   50
               TabIndex        =   22
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   540
            TabIndex        =   31
            Top             =   780
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "DEPARTMENT"
            Height          =   195
            Left            =   540
            TabIndex        =   30
            Top             =   420
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PERIODE"
            Height          =   195
            Left            =   540
            TabIndex        =   29
            Top             =   1560
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "MONTH"
            Height          =   195
            Left            =   540
            TabIndex        =   28
            Top             =   1170
            Width           =   600
         End
      End
      Begin VB.Frame fra_monthly 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   4
         Top             =   660
         Width           =   8415
         Begin VB.ComboBox cbo_monthly_company 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":6855
            Left            =   1800
            List            =   "frm_rpt_salary_summary.frx":685F
            TabIndex        =   10
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.Frame fra_monthly_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   6
            Top             =   960
            Width           =   4575
            Begin VB.CommandButton cmd_monthly_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1440
               TabIndex        =   9
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
               TabIndex        =   8
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
               TabIndex        =   7
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.ComboBox cbo_monthly_employee 
            Height          =   315
            ItemData        =   "frm_rpt_salary_summary.frx":6872
            Left            =   1800
            List            =   "frm_rpt_salary_summary.frx":687C
            TabIndex        =   5
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_monthly 
            Height          =   300
            Left            =   1800
            TabIndex        =   11
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   92078083
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Perusahaan"
            Height          =   195
            Left            =   720
            TabIndex        =   14
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Bulan"
            Height          =   195
            Left            =   720
            TabIndex        =   13
            Top             =   1620
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Karyawan"
            Height          =   195
            Left            =   720
            TabIndex        =   12
            Top             =   1200
            Width           =   705
         End
      End
      Begin VB.Frame fra_process_monthly 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   15
         Top             =   660
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
            TabIndex        =   16
            Top             =   1080
            Width           =   2730
         End
      End
      Begin VB.Frame fra_process_periode 
         Height          =   2475
         Left            =   780
         TabIndex        =   42
         Top             =   1020
         Visible         =   0   'False
         Width           =   8415
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   1770
            Width           =   8205
            _ExtentX        =   14473
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label13 
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
            TabIndex        =   33
            Top             =   270
            Width           =   2730
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            Height          =   405
            Left            =   3660
            TabIndex        =   45
            Top             =   1230
            Width           =   4575
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            Height          =   345
            Left            =   450
            TabIndex        =   44
            Top             =   1230
            Width           =   2385
         End
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1230
      OleObjectBlob   =   "frm_rpt_salary_summary.frx":688D
      TabIndex        =   2
      Top             =   990
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1080
      Top             =   990
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
      Left            =   9060
      Top             =   780
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1050
      Width           =   795
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT SALARY"
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
      TabIndex        =   32
      Top             =   150
      Width           =   4365
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_rpt_salary_summary.frx":87F3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14850
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
Private WithEvents poSendMail As clsSendMail
Attribute poSendMail.VB_VarHelpID = -1

' misc local vars
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean

Dim int_libur() As Integer
Dim pub_employee_code, pub_employee_name, pub_company_code, pub_company_name, pub_email As String
Dim int_sent_false, int_sent_true As Integer
Dim str_smtp, str_sender_mail, str_sender_name As String
Dim rs_employee As New ADODB.Recordset

Private Sub cbo_periode_employee_Click()
    If cbo_periode_employee.ListIndex = 0 Then
        fra_periode_employee.Visible = False
    Else
        fra_periode_employee.Visible = True
    End If
    
    txt_periode_employee_code = "": txt_periode_employee_name = "": txt_periode_nik = ""
End Sub

'Private Sub cmd_periode_browse_employee_Click()
'    frm_lookup_mst_employee.public_int_mode = 79
'    frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
'    frm_lookup_mst_employee.Show 1
'End Sub

Private Sub cmd_print_slip_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
    If SSTab1.Tab = 1 Then
        If check_valid_periode Then
            Call rpt_periode(1)
        End If
    End If
End Sub

Private Sub cmd_send_mail_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
    int_sent_false = 0
    int_sent_true = 0
    
    If SSTab1.Tab = 1 Then
        If check_valid_periode Then
            Call send_mail_periode(1)
        End If
    End If
End Sub

Private Sub send_mail_periode(ByVal j As Integer)
Dim str_sql, str_param_periode, str_file, str1, str_file_out As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt

    If j = 0 Then
        str_file = "\report\rpt_53_1.rpt"
    ElseIf j = 1 Then
        str_file = "\report\rpt_53.rpt"
    End If
    
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
        str1 = "select * from m_employee where ifnull(flag_active,0)=1 order by employee_code asc"
        
    ElseIf int_flag_company = 1 Then
        If cbo_periode_employee.ListIndex = 0 Then
            int_flag_employee = 0
            str_employee_code = "-"
            str1 = "select * from m_employee where company_code='" & str_company_code & "' and ifnull(flag_active,0)=1 order by employee_code asc"
        ElseIf cbo_periode_employee.ListIndex = 1 Then
            int_flag_employee = 1
            str_employee_code = txt_periode_employee_code
            str1 = "select * from m_employee where employee_code='" & str_employee_code & "' and ifnull(flag_active,0)=1 order by employee_code asc"
        End If
    End If
    
    If rs_employee.State = 1 Then rs_employee.Close
    rs_employee.Open str1, CnG, adOpenStatic, adLockReadOnly
    If rs_employee.RecordCount > 0 Then rs_employee.MoveFirst
    
    While Not rs_employee.EOF
    
        If cbo_periode_to.ListIndex = 0 Then
            str_sql = "call spr_attendance_54('" _
                & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" _
                & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "'," _
                & 1 & ",'" & rs_employee!COMPANY_CODE & "'," & 1 & ",'" & rs_employee!employee_code & "')"
            str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
        ElseIf cbo_periode_to.ListIndex = 1 Then
            str_sql = "call spr_attendance_54('" _
                & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" _
                & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "'," _
                & 1 & ",'" & rs_employee!COMPANY_CODE & "'," & 1 & ",'" & rs_employee!employee_code & "')"
            str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
                & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
        End If
    
        
        '-- creating pdf
        str_file_out = App.Path & "\mail\slip_" & Format(DTPicker_periode_from, "yyyymm") & "_" & rs_employee!employee_code & ".pdf"
        Call rpt_auto_pdf(str_sql, str_file, str_file_out, str_param_periode)
        
        pub_employee_code = rs_employee!employee_code
        pub_employee_name = rs_employee!EMPLOYEE_NAME
        pub_company_code = rs_employee!COMPANY_CODE
        pub_company_name = rs_employee!company_name
        pub_email = "" & rs_employee!email
        
        Call send_mail("Salary slip " & str_param_periode, "Salary slip " & str_param_periode & vbCrLf & vbCrLf _
                        & "TTD" & vbCrLf & str_sender_name, str_file_out)
        '--
    
        rs_employee.MoveNext
    Wend
    
    Call show_msg
End Sub

Private Sub show_msg()
    MsgBox "There are " & int_sent_true & " mail are sent successfully!" & vbCrLf _
        & int_sent_false & " are fails!" & vbCrLf _
        & "For more detail info, let see the 'log mail'!", vbInformation, headerMSG
End Sub

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

Private Sub rpt_periode(ByVal j As Integer)
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer
Dim strsql As String
Dim rsEmployee As New ADODB.Recordset

    int_process = vbNo
    
    If j = 0 Then
        str_file = "\report\rpt_summary_salary.rpt"
    ElseIf j = 2 Then
        str_file = "\report\rpt_summary_salary.rpt"
    ElseIf j = 1 Then
        str_file = "\report\rpt_slip_salary.rpt"
    End If
    
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
    
        If j = 0 Or j = 1 Then
            Call getPeriodeProses
    
            strsql = "SELECT 1 FROM h_salary a join m_employee b on a.employee_code = b.employee_code WHERE `month` = '" & Format(DTPicker1.Value, "yyyy-MM") & "' " _
                & "AND a.company_code = '" & TDBCombo_company.Text & "'"
            rsEmployee.Open strsql, CnG, adOpenStatic, adLockReadOnly
            
            If rsEmployee.RecordCount = 0 Then
                MsgBox "Salary Data For This Department Has Not Been Calculate...!!"
                rsEmployee.Close
                Exit Sub
            End If
            rsEmployee.Close
        End If
    
    'If Not j = 2 Then
    '    int_process = MsgBox("Would you want to process data before run report?", vbYesNo, headerMSG)
    'End If
    d1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
    d2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
    
        If j = 0 Then
        '--
            If cbo_periode_to.ListIndex = 0 Then
                str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                    & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
                    & int_flag_employee & ",'" & str_employee_code & "','" & LOGIN_LEVEL & "')"
                str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM") & ")"
            ElseIf cbo_periode_to.ListIndex = 1 Then
                str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                    & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
                    & int_flag_employee & ",'" & str_employee_code & "','" & LOGIN_LEVEL & "')"
                str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM") & ")"
            End If
        '--
        ElseIf j = 1 Then
        '--
            If cbo_periode_to.ListIndex = 0 Then
                str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                    & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
                    & int_flag_employee & ",'" & str_employee_code & "','" & LOGIN_LEVEL & "')"
                str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM") & ")"
                'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
            ElseIf cbo_periode_to.ListIndex = 1 Then
                str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                    & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
                    & int_flag_employee & ",'" & str_employee_code & "','" & LOGIN_LEVEL & "')"
                str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM") & ")"
                'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
                    & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
            End If
            '--
        ElseIf j = 2 Then
            str_sql = "call spr_salary_hardys_est('" & d1 & "','" & d2 & "'," _
                & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
                & int_flag_employee & ",'" & str_employee_code & "')"
        End If
    
    
    Text1 = str_sql
    
    Call a.Show
    a.Caption = "SALARY REPORT"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
    
    fra_process_periode.Visible = False
    fra_periode.Visible = True
End Sub

Private Sub cmdPrint_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If

    If SSTab1.Tab = 1 Then
        If check_valid_periode Then
            Call rpt_periode(0)
        End If
    End If
End Sub

Private Sub DTPicker1_Change()
    Call getPeriodeProses
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    Call getPeriodeProses
End Sub

Private Sub Form_Load()
    Adodc_company.ConnectionString = strConn
    Adodc_department.ConnectionString = strConn
    
    Label10.Caption = ""
    Label11.Caption = ""
    Label10.Visible = False
    Label11.Visible = False
    ProgressBar1.Visible = False
    fra_process_periode.Visible = False
    
    Call load_data_company
'    Call load_data_setting_mail
    
    DTPicker_periode_from.Value = Now
    DTPicker_periode_to.Value = Now
    DTPicker_monthly.Value = Now
    DTPicker1.Value = Now
    
    Call getPeriodeProses
    
    cbo_periode_to.ListIndex = 0
    cbo_periode_company.ListIndex = 1
    cbo_periode_employee.ListIndex = 0
    
    cbo_monthly_company.ListIndex = 1
    cbo_monthly_employee.ListIndex = 0
    
    Timer1.Enabled = True
    SSTab1.TabVisible(0) = False
    
    cbo_periode_to.ListIndex = 1
    
    Call createGridKar
End Sub

Private Sub load_data_company()
    Adodc_company.RecordSource = "select * from m_company order by company_code"
    Adodc_company.Refresh
    
    TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If
    
    Call load_data_department
End Sub


Private Sub Timer1_Timer()
Timer1.Enabled = False
    Call set_company_mode_adodc(Adodc_company, TDBCombo_company, txt_company_name)
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

Private Sub load_data_setting_mail()
'Dim rs1 As New ADODB.Recordset
'
'    rs1.Open "select * from s_mail where s_number=1", CnG, adOpenStatic, adLockReadOnly
'    If rs1.RecordCount > 0 Then
'        str_smtp = rs1.Fields("s_smtp").Value
'        str_sender_mail = rs1.Fields("s_sender_email").Value
'        str_sender_name = rs1.Fields("s_sender_name").Value
'    Else
'        str_smtp = ""
'        str_sender_mail = ""
'        str_sender_name = ""
'        MsgBox "No valid SMTP setting!", vbInformation, headerMSG
'        'Unload Me
'    End If
End Sub

Private Sub send_mail(ByVal str_subject As String, ByVal str_msg As String, ByVal str_attc As String)
Set poSendMail = New clsSendMail

    With poSendMail
    
        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************
    
        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)
    
        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = str_smtp
        .from = str_sender_mail
        .FromDisplayName = str_sender_name
        .Recipient = rs_employee.Fields("email").Value
        .RecipientDisplayName = rs_employee.Fields("employee_nick_name").Value
    '        .CcRecipient = "CcRecipient"
    '        .CcDisplayName = "CcDisplayName"
    '        .BccRecipient = "BccRecipient"
        '.ReplyToAddress = txtFrom.Text              ' Optional, used when different than 'From' address
        .Subject = str_subject                  ' Optional
        .Message = str_msg                      ' Optional
        .Attachment = str_attc          ' Optional, separate multiple entries with delimiter character
    
        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
    '        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
    '        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
    '        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
    '        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
    '        .Receipt = bReceipt                         ' Optional, default = FALSE
    '        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
    '        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
    '        .Username = txtUserName                     ' Optional, default = Null String
    '        .Password = txtPassword                     ' Optional, default = Null String, value is NOT saved
    '        .POP3Host = txtPopServer
    '        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
        ' .SMTPPort = 25                            ' Optional, default = 25
    
        ' **************************************************************************
        ' OK, all of the properties are set, send the email...
        ' **************************************************************************
        ' .Connect                                  ' Optional, use when sending bulk mail
        .Send                                       ' Required
        ' .Disconnect                               ' Optional, use when sending bulk mail
    '        txtServer.Text = .SMTPHost                  ' Optional, re-populate the Host in case
                                                    ' MX look up was used to find a host    End With
    End With
    End Sub

Private Sub poSendMail_SendFailed(Explanation As String)
Dim rs1 As New ADODB.Recordset

    rs1.Open "select * from h_send_mail where employee_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic
    
    CnG.BeginTrans
    
    With rs1
        .AddNew
        .Fields("date").Value = Now
        .Fields("employee_code").Value = pub_employee_code
        .Fields("employee_name").Value = pub_employee_name
        .Fields("email").Value = pub_email
        .Fields("sent_status").Value = 0
        .Fields("description").Value = "Your attempt to send mail failed for the following reason(s): " _
                                        & vbCrLf & Explanation
        .Update
    End With
    
    CnG.CommitTrans
    
    int_sent_false = int_sent_false + 1
End Sub

Private Sub poSendMail_SendSuccesful()
Dim rs1 As New ADODB.Recordset

    rs1.Open "select * from h_send_mail where employee_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic
    
    CnG.BeginTrans
    
    With rs1
        .AddNew
        .Fields("date").Value = Now
        .Fields("employee_code").Value = pub_employee_code
        .Fields("employee_name").Value = pub_employee_name
        .Fields("email").Value = pub_email
        .Fields("sent_status").Value = 1
        .Fields("description").Value = "sent successfully"
        .Update
    End With
    
    CnG.CommitTrans
    
    int_sent_true = int_sent_true + 1
End Sub

Private Sub set_data_department(ByVal str_code As String)
On Error Resume Next

    Adodc_department.Recordset.MoveFirst
    Adodc_department.Recordset.Find ("department_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (Adodc_department.Recordset.EOF = True Or Adodc_department.Recordset.BOF = True) Then
        TDBCombo_department.Bookmark = Adodc_department.Recordset.AbsolutePosition
        Call tdbCombo_department_itemChange
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

Private Sub tdbCombo_department_itemChange()
    If TDBCombo_department.ApproxCount > 0 Then
        TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
        txt_department_name = TDBCombo_department.Columns("department_name").Value
    End If
End Sub

Private Sub getPeriodeProses()
Dim strsql As String
Dim rsbulan As New ADODB.Recordset

    strsql = "select periode_from,periode_to from h_d_salary " _
            & "WHERE company_code = '" & TDBCombo_company.Text & "' " _
            & "AND left(`month`,7) = '" & Format(DTPicker1.Value, "yyyy-MM") & "'"
    rsbulan.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsbulan.RecordCount > 0 Then
        DTPicker_periode_from.Value = rsbulan!periode_from
        DTPicker_periode_to.Value = rsbulan!periode_to
        cbo_periode_to.ListIndex = 1
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
Dim vParam As String
    If pilihan = 1 Then
        LynxGrid2.Clear
        
        vParam = IIf(DEPARTMENT_CODE <> "" And DIVISION_CODE = "", "a.department_code = '" & DEPARTMENT_CODE & "'", IIf(DEPARTMENT_CODE = "" And DIVISION_CODE = "", "a.company_code = '" & COMPANY_CODE & "'", "a.department_code = '" & DEPARTMENT_CODE & "' AND a.division_code = '" & DIVISION_CODE & "'"))
        
        If LOGIN_LEVEL = 100 Then
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND (nik LIKE '%" & txt_periode_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_periode_nik.Text & "%')"
        Else
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND " & vParam & " " & _
                        "AND (nik LIKE '%" & txt_periode_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_periode_nik.Text & "%') " & _
                        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
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
                txt_periode_employee_code.Text = rs!employee_code
                txt_periode_employee_name.Text = rs!EMPLOYEE_NAME
                txt_periode_nik.Text = rs!nik
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
            txt_periode_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txt_periode_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txt_periode_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
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

Private Sub txt_periode_nik_Change()
    If txt_periode_nik.Text = "" Then
        txt_periode_employee_code.Text = ""
        txt_periode_employee_name.Text = ""
    End If
End Sub

Private Sub txt_periode_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmd_periode_browse_employee_Click()
    isiGridKar (1)
End Sub
