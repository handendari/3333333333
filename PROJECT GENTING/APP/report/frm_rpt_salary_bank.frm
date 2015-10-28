VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_rpt_salary_bank 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " SALARY BANK REPORT"
   ClientHeight    =   6825
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   10560
   Icon            =   "frm_rpt_salary_bank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Text            =   "frm_rpt_salary_bank.frx":000C
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   600
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "MONTHLY"
      TabPicture(0)   =   "frm_rpt_salary_bank.frx":0012
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_process_monthly"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_monthly"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "PERIODE"
      TabPicture(1)   =   "frm_rpt_salary_bank.frx":002E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_periode"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra_process_periode"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "RESERVED"
      TabPicture(2)   =   "frm_rpt_salary_bank.frx":004A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame fra_monthly 
         Height          =   2655
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Width           =   8415
         Begin VB.ComboBox cbo_monthly_company 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":0066
            Left            =   1800
            List            =   "frm_rpt_salary_bank.frx":0070
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
            ItemData        =   "frm_rpt_salary_bank.frx":0083
            Left            =   1800
            List            =   "frm_rpt_salary_bank.frx":008D
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
            Format          =   164102147
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label13 
            Caption         =   "* yyyy-MM"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   3540
            TabIndex        =   53
            Top             =   1590
            Width           =   1125
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
            Top             =   1560
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
         Left            =   840
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
         Height          =   2925
         Left            =   -74160
         TabIndex        =   18
         Top             =   960
         Width           =   8415
         Begin VB.TextBox txtAuthorized 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1800
            TabIndex        =   49
            Top             =   2370
            Width           =   3465
         End
         Begin VB.TextBox txtPrepare 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1800
            TabIndex        =   48
            Top             =   1980
            Width           =   3465
         End
         Begin VB.TextBox txt_periode_employee_code 
            Height          =   285
            Left            =   4590
            TabIndex        =   47
            Text            =   "Text2"
            Top             =   180
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Frame fra_periode_department 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   3600
            TabIndex        =   44
            Top             =   420
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
               TabIndex        =   45
               Top             =   90
               Width           =   2415
            End
            Begin TrueOleDBList60.TDBCombo TDBCombo_department 
               Height          =   375
               Left            =   0
               OleObjectBlob   =   "frm_rpt_salary_bank.frx":009E
               TabIndex        =   46
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
         Begin VB.ComboBox cbo_periode_department 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":205F
            Left            =   1800
            List            =   "frm_rpt_salary_bank.frx":2069
            TabIndex        =   43
            Text            =   "..."
            Top             =   510
            Width           =   1695
         End
         Begin VB.ComboBox cbo_bank_name 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":207C
            Left            =   1800
            List            =   "frm_rpt_salary_bank.frx":207E
            TabIndex        =   41
            Text            =   "..."
            Top             =   1620
            Width           =   1695
         End
         Begin VB.ComboBox cbo_periode_to 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":2080
            Left            =   3600
            List            =   "frm_rpt_salary_bank.frx":208A
            TabIndex        =   26
            Text            =   "..."
            Top             =   1260
            Width           =   1335
         End
         Begin VB.ComboBox cbo_periode_company 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":2098
            Left            =   3570
            List            =   "frm_rpt_salary_bank.frx":20A2
            TabIndex        =   25
            Text            =   "..."
            Top             =   1620
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "DAY COUNT"
            Height          =   495
            Left            =   0
            TabIndex        =   24
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cbo_periode_employee 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":20B5
            Left            =   1800
            List            =   "frm_rpt_salary_bank.frx":20BF
            TabIndex        =   23
            Text            =   "..."
            Top             =   900
            Width           =   1695
         End
         Begin VB.Frame fra_periode_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   585
            Left            =   3600
            TabIndex        =   19
            Top             =   660
            Width           =   4575
            Begin VB.TextBox txt_periode_nik 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   0
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   22
               Top             =   240
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
               Top             =   240
               Width           =   2415
            End
            Begin VB.CommandButton cmd_periode_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1440
               TabIndex        =   20
               Top             =   240
               Width           =   375
            End
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_from 
            Height          =   300
            Left            =   1800
            TabIndex        =   27
            Top             =   1260
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   161021955
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_to 
            Height          =   300
            Left            =   5040
            TabIndex        =   28
            Top             =   1260
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   161021955
            CurrentDate     =   39278
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Authorized By"
            Height          =   195
            Left            =   720
            TabIndex        =   51
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Prepared By"
            Height          =   195
            Left            =   720
            TabIndex        =   50
            Top             =   2040
            Width           =   870
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Bank Name"
            Height          =   195
            Left            =   720
            TabIndex        =   42
            Top             =   1620
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Periode"
            Height          =   195
            Left            =   720
            TabIndex        =   31
            Top             =   1260
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Department"
            Height          =   195
            Left            =   720
            TabIndex        =   30
            Top             =   540
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Employee"
            Height          =   195
            Left            =   720
            TabIndex        =   29
            Top             =   900
            Width           =   690
         End
      End
      Begin VB.Frame fra_process_periode 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   35
         Top             =   960
         Width           =   8415
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
            Left            =   2640
            TabIndex        =   36
            Top             =   1080
            Width           =   2730
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   5460
      Width           =   10095
      Begin VB.CommandButton cmd_print_bank 
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
         Left            =   4800
         Picture         =   "frm_rpt_salary_bank.frx":20D0
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   360
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
         Left            =   1800
         Picture         =   "frm_rpt_salary_bank.frx":265A
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Summary"
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
         Left            =   2880
         Picture         =   "frm_rpt_salary_bank.frx":2BE4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
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
         Picture         =   "frm_rpt_salary_bank.frx":316E
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
         Left            =   6600
         Picture         =   "frm_rpt_salary_bank.frx":36F8
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
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
         Left            =   8280
         Picture         =   "frm_rpt_salary_bank.frx":3C82
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_rpt_salary_bank.frx":420C
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1080
      Top             =   600
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
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "LAPORAN SALARY BANK"
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
      Left            =   3510
      TabIndex        =   52
      Top             =   0
      Width           =   3945
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   795
   End
End
Attribute VB_Name = "frm_rpt_salary_bank"
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


Private Sub report_kehadiran_karyawan()
'Dim lstr_criteria As String
'
'CrystalReport1.Reset
'CrystalReport1.ReportFileName = App.Path & "\report\rpt_kehadiran_karyawan.rpt"
'
'lstr_criteria = "({v_att_karyawan.flag_io}) = 0 and Date({v_att_karyawan.tanggal}) in #" _
'                & Format(DTPicker_bulan.Value, "yyyy,mm,01") & "# to #" _
'                & Format(DTPicker_bulan.Value, "yyyy,mm,31") & "#"
'
'CrystalReport1.ParameterFields(0) = "p_karyawan;Karyawan : (" _
'& TDBCombo_karyawan.Columns("kode_karyawan").Value & ") " _
'& TDBCombo_karyawan.Columns("nama_karyawan").Value & ";true"
'
'CrystalReport1.ReplaceSelectionFormula (lstr_criteria)
'
'CrystalReport1.WindowState = crptMaximized
'CrystalReport1.Action = 1
End Sub

Private Sub cbo_daily_company_Click()
'If cbo_daily_company.ListIndex = 0 Then
'    cbo_daily_employee.ListIndex = 0
'    cbo_daily_employee.Enabled = False
'ElseIf cbo_daily_company.ListIndex = 1 Then
'    cbo_daily_employee.ListIndex = 0
'    cbo_daily_employee.Enabled = True
'End If
End Sub

Private Sub cbo_daily_employee_Click()
'If cbo_daily_employee.ListIndex = 0 Then
'    fra_daily_employee.Visible = False
'Else
'    fra_daily_employee.Visible = True
'    txt_daily_employee_code = "": txt_daily_employee_name = ""
'End If
End Sub

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

Private Sub cbo_periode_company_Click()
If cbo_periode_company.ListIndex = 0 Then
    cbo_periode_employee.ListIndex = 0
    cbo_periode_employee.Enabled = False
ElseIf cbo_periode_company.ListIndex = 1 Then
    cbo_periode_employee.ListIndex = 0
    cbo_periode_employee.Enabled = True
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
End If

    txt_periode_employee_code = "": txt_periode_employee_name = "": txt_periode_nik = ""
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

Private Sub cmd_print_bank_Click()
If check_validate_tdbcombo(TDBCombo_company) = False Then
    MsgBox "No Company selected!", vbInformation, headerMSG
    Exit Sub
End If

'If SSTab1.Tab = 1 Then
'    If check_valid_periode Then
'        Call rpt_bank(cbo_bank_name.ListIndex)
'    End If
'End If

If SSTab1.Tab = 1 Then
'    If check_valid_periode Then
'        If cbo_periode_employee.Text = "All" And cbo_bank_name.Text = "DANAMON" Then
'            Call rpt_bank(0)
'        ElseIf cbo_periode_employee.Text = "Select" And cbo_bank_name.Text = "DANAMON" Then
'            Call rpt_bank(1)
'        ElseIf cbo_periode_employee.Text = "All" And cbo_bank_name.Text = "MANDIRI" Then
'            Call rpt_bank(2)
'        ElseIf cbo_periode_employee.Text = "Select" And cbo_bank_name.Text = "MANDIRI" Then
'            Call rpt_bank(3)
'        End If
'    End If
    Call rpt_bank
End If
End Sub

Private Sub cmd_print_slip_Click()
If check_validate_tdbcombo(TDBCombo_company) = False Then
    MsgBox "No Company selected!", vbInformation, headerMSG
    Exit Sub
End If

If SSTab1.Tab = 1 Then
    If check_valid_periode Then
        Call rpt_periode(1)
    End If
ElseIf SSTab1.Tab = 0 Then
    If check_valid_monthly Then
        Call rpt_monthly(1)
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
ElseIf SSTab1.Tab = 0 Then
    If check_valid_monthly Then
        Call send_mail_monthly(1)
    End If

End If
End Sub

Private Sub send_mail_monthly(ByVal j As Integer)
Dim str_sql, str_param_periode, str_file, str1, str_file_out As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim dt1, dt2 As Date
Dim d1, d2 As String
Dim a As New frm_rpt

If j = 0 Then
    str_file = "\report\rpt_53_1.rpt"
ElseIf j = 1 Then
    str_file = "\report\rpt_53.rpt"
End If

If cbo_monthly_company.ListIndex = 0 Then
    int_flag_company = 0
    str_company_code = "-"
ElseIf cbo_monthly_company.ListIndex = 1 Then
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
End If


If int_flag_company = 0 Then
    int_flag_employee = 0
    str_employee_code = "-"
    str1 = "select * from m_employee where ifnull(flag_active,0)=1 order by employee_code asc"
ElseIf int_flag_company = 1 Then
    If cbo_monthly_employee.ListIndex = 0 Then
        int_flag_employee = 0
        str_employee_code = "-"
        str1 = "select * from m_employee where company_code='" & str_company_code & "' and ifnull(flag_active,0)=1 order by employee_code asc"
    ElseIf cbo_monthly_employee.ListIndex = 1 Then
        int_flag_employee = 1
        str_employee_code = txt_monthly_employee_code
        str1 = "select * from m_employee where employee_code='" & str_employee_code & "' and ifnull(flag_active,0)=1 order by employee_code asc"
    End If
End If

If rs_employee.State = 1 Then rs_employee.Close
rs_employee.Open str1, CnG, adOpenStatic, adLockReadOnly
If rs_employee.RecordCount > 0 Then rs_employee.MoveFirst

dt1 = Format(DTPicker_monthly.Value, "yyyy-MM-01")
dt2 = DateAdd("m", 1, dt1): dt2 = DateAdd("d", -1, dt2)
d1 = Format(dt1, "yyyy-mm-dd")
d2 = Format(dt2, "yyyy-mm-dd")

While Not rs_employee.EOF

    str_sql = "call spr_attendance_54('" & d1 & "','" & d2 & "'," _
            & 1 & ",'" & rs_employee!COMPANY_CODE & "'," & 1 & ",'" & rs_employee!employee_code & "')"
    str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"

    '-- creating pdf
    str_file_out = App.Path & "\mail\slip_" & Format(DTPicker_monthly, "yyyymm") & "_" & rs_employee!employee_code & ".pdf"
    Call rpt_auto_pdf(str_sql, str_file, str_file_out, str_param_periode)
    
    pub_employee_code = rs_employee!employee_code
    pub_employee_name = rs_employee!employee_name
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
    pub_employee_name = rs_employee!employee_name
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

'Private Sub cbo_department_Click()
'If cbo_department.Text = "All" Then
'    TDBCombo_department.Visible = False
'    txt_department_name.Visible = False
'Else
'    TDBCombo_department.Visible = True
'    txt_department_name.Visible = True
'    TDBCombo_department.Text = ""
'    txt_department_name.Text = ""
'
'    Call load_data_department
'End If
'End Sub

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

Private Sub Command2_Click()
'MsgBox UBound(int_libur)

'Call get_holiday("2007-07")

Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

cmd.ActiveConnection = CnG
cmd.CommandText = Text1.Text
rs.CursorLocation = adUseClient
rs.Open cmd, , adOpenStatic, adLockReadOnly

MsgBox rs.RecordCount
End Sub


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

Private Sub rpt_bank()
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer

int_process = vbNo
'If j = 0 Then
'    str_file = "\report\rpt_danamon.rpt"
'ElseIf j = 1 Then
'    str_file = "\report\rpt_danamon.rpt"
'ElseIf j = 2 Then
'    str_file = "\report\rpt_mandiri.rpt"
'ElseIf j = 3 Then
'    str_file = "\report\rpt_mandiri.rpt"
'End If

str_file = "\report\rpt_bank.rpt"

int_flag_company = 1
str_company_code = TDBCombo_company.Columns("company_code").Value
int_flag_employee = 0

d1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
d2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")

If cbo_periode_employee.Text = "All" Then
    str_employee_code = ""
Else
    str_employee_code = txt_periode_employee_code.Text
End If

'If cbo_department.Text = "All" Then
'    str_sql = "call spr_salary_hardys_bank_all('" & d1 & "','" & d2 & "'," _
'        & int_flag_company & ",'" & str_company_code & "'," _
'        & int_flag_employee & ",'" & str_employee_code & "','" & cbo_bank_name.Text & "')"
'Else
str_sql = "call spr_salary_hardys_bank('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
        & int_flag_employee & ",'" & str_employee_code & "','" & cbo_bank_name.Text & "'," _
        & cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "','" & txtPrepare.Text & "','" & txtAuthorized.Text & "')"
'End If

Text1 = str_sql

Call a.Show
a.Caption = "SALARY BANK REPORT"
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub



Private Sub rpt_periode(ByVal j As Integer)
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer

int_process = vbNo

If j = 0 Then
    str_file = "\report\rpt_hardys_salary1.rpt"
ElseIf j = 2 Then
    str_file = "\report\rpt_hardys_salary1.rpt"
ElseIf j = 1 Then
    str_file = "\report\rpt_hardys_salary2.rpt"
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


If Not j = 2 Then
    int_process = MsgBox("Would you want to process data before run report?", vbYesNo, headerMSG)
End If
d1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
d2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")

If int_process = vbYes Then
    If int_flag_employee = 1 Then
        str1 = "DELETE FROM h_salary WHERE employee_code = '" _
                & str_employee_code & "' AND LEFT(MONTH,7) = '" & Format(d1, "yyyy-MM") & "';"
        str2 = "CALL spg_salary('" & str_company_code & "', 1, '" & str_employee_code & "', " _
                & "'" & d1 & "', '" & d2 & "', '" & d2 & "');"
    Else
        str1 = "DELETE FROM h_salary WHERE LEFT(MONTH,7) = '" & Format(d1, "yyyy-MM") & "';"
        str2 = "CALL spg_salary('" & str_company_code & "', 0, '" & str_employee_code & "', " _
                & "'" & d1 & "', '" & d2 & "', '" & d2 & "');"
    End If
    
    CnG.Execute str1
    CnG.Execute str2
End If


If j = 0 Or j = 2 Then
'--
If cbo_periode_to.ListIndex = 0 Then
    str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
ElseIf cbo_periode_to.ListIndex = 1 Then
    str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
        & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
End If
'--
ElseIf j = 1 Then
'--
If cbo_periode_to.ListIndex = 0 Then
    str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
ElseIf cbo_periode_to.ListIndex = 1 Then
    str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
        & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
End If
'--
End If


Text1 = str_sql

Call a.Show
a.Caption = "SALARY REPORT"
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub rpt_monthly(ByVal j As Integer)
Dim str_sql, str_param_periode, str_file, str1, str2 As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer

int_process = vbNo

If j = 0 Then
    str_file = "\report\rpt_53_1_ebiz.rpt"
ElseIf j = 2 Then
    str_file = "\report\rpt_53_1_ebiz_list.rpt"
ElseIf j = 1 Then
    str_file = "\report\rpt_53_ebiz.rpt"
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

If Not j = 2 Then
    int_process = MsgBox("Would you want to process data before run report?", vbYesNo, headerMSG)
End If
d1 = Format(DTPicker_monthly.Value, "yyyy-MM-01"): dx = DateAdd("m", 1, d1)
d2 = Format(d1, "yyyy-MM-") & Format(DateDiff("d", d1, dx), "0#")

If int_process = vbYes Then
    If int_flag_employee = 1 Then
        str1 = "DELETE FROM h_salary WHERE employee_code = '" _
                & str_employee_code & "' AND LEFT(MONTH,7) = '" & Format(d1, "yyyy-MM") & "';"
        str2 = "CALL spg_salary('" & str_company_code & "', 1, '" & str_employee_code & "', " _
                & "'" & d1 & "', '" & d2 & "', '" & d2 & "');"
    Else
        str1 = "DELETE FROM h_salary WHERE LEFT(MONTH,7) = '" & Format(d1, "yyyy-MM") & "';"
        str2 = "CALL spg_salary('" & str_company_code & "', 0, '" & str_employee_code & "', " _
                & "'" & d1 & "', '" & d2 & "', '" & d2 & "');"
    End If
    
    CnG.Execute str1
    CnG.Execute str2
End If

If j = 0 Or j = 2 Then
    '--
    str_sql = "call spr_salary_ebiz_sum('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    'str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
    '--
ElseIf j = 1 Then
    '--
    str_sql = "call spr_salary_ebiz_sli('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    'str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
    '--
End If

Text1 = str_sql

Call a.Show
a.Caption = "SALARY REPORT"
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub CmdPrint_Click()
If check_validate_tdbcombo(TDBCombo_company) = False Then
    MsgBox "No Company selected!", vbInformation, headerMSG
    Exit Sub
End If

If SSTab1.Tab = 1 Then
    If check_valid_periode Then
        Call rpt_periode(0)
    End If
ElseIf SSTab1.Tab = 0 Then
    If check_valid_monthly Then
        Call rpt_monthly(0)
    End If

End If
End Sub

Private Sub cmd_print_sum_list_Click()
If check_validate_tdbcombo(TDBCombo_company) = False Then
    MsgBox "No Company selected!", vbInformation, headerMSG
    Exit Sub
End If

If SSTab1.Tab = 1 Then
    If check_valid_periode Then
        Call rpt_periode(2)
    End If
ElseIf SSTab1.Tab = 0 Then
    If check_valid_monthly Then
        Call rpt_monthly(2)
    End If

End If
End Sub

Private Sub Command3_Click()
Adodc1.ConnectionString = strConn
Adodc1.RecordSource = Text1.Text
Adodc1.Refresh

MsgBox Adodc1.Recordset.RecordCount
End Sub

Private Sub Form_Load()
Adodc_company.ConnectionString = strConn
Adodc_department.ConnectionString = strConn

Call load_data_company
'Call load_data_setting_mail
Call load_data_bank
Call load_data_user_access(Me)

DTPicker_periode_from.Value = Now
DTPicker_periode_to.Value = Now
DTPicker_monthly.Value = Now

cbo_periode_to.ListIndex = 0
cbo_periode_company.ListIndex = 1
cbo_periode_employee.ListIndex = 0
'cbo_department.Text = "All"

'TDBCombo_department.Visible = False
'txt_department_name.Visible = False

cbo_monthly_company.ListIndex = 1
cbo_monthly_employee.ListIndex = 0

txtPrepare = ""
txtAuthorized = ""

Timer1.Enabled = True
SSTab1.TabVisible(0) = False
cbo_bank_name.ListIndex = 0
cbo_periode_department.ListIndex = 0
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

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
End If

Call load_data_department
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
Timer1.Enabled = False
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
'
'rs1.Open "select * from s_mail where s_number=1", CnG, adOpenStatic, adLockReadOnly
'If rs1.RecordCount > 0 Then
'    str_smtp = rs1.Fields("s_smtp").Value
'    str_sender_mail = rs1.Fields("s_sender_email").Value
'    str_sender_name = rs1.Fields("s_sender_name").Value
'Else
'    str_smtp = ""
'    str_sender_mail = ""
'    str_sender_name = ""
'    MsgBox "No valid SMTP setting!", vbInformation, headerMSG
'    'Unload Me
'End If
'End Sub

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
    Call TDBCombo_department_ItemChange
Else
    TDBCombo_department.Text = ""
End If
End Sub

Private Sub load_data_department()
TDBCombo_department.Text = "": txt_department_name = ""

Adodc_department.RecordSource = "select department_code, department_name from m_department where company_code='" _
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

Private Sub poSendMail_Status(Status As String)
'    vbSendMail 'Status Event'
'    lstStatus.AddItem Status
'    lstStatus.ListIndex = lstStatus.ListCount - 1
'    lstStatus.ListIndex = -1
End Sub

Private Sub poSendMail_Progress(lPercentCompete As Long)
'   vbSendMail 'Progress Event'
'   lblProgress = lPercentCompete & "% complete"
End Sub

Private Sub load_data_bank()
Dim strsql As String
Dim rs As New ADODB.Recordset

strsql = "SELECT bank_code from m_bank order by bank_code"
rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then

If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        cbo_bank_name.AddItem rs!bank_code
    rs.MoveNext
    Wend
End If

End If
rs.Close

End Sub
