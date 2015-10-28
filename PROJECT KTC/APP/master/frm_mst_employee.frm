VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_employee 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER EMPLOYEE"
   ClientHeight    =   9225
   ClientLeft      =   -15
   ClientTop       =   345
   ClientWidth     =   14685
   Icon            =   "frm_mst_employee.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "List Employee Status"
      Height          =   585
      Left            =   11490
      TabIndex        =   104
      Top             =   60
      Width           =   2925
      Begin VB.OptionButton optNotActive 
         Caption         =   "Not Active"
         Height          =   225
         Left            =   1500
         TabIndex        =   106
         Top             =   270
         Width           =   1245
      End
      Begin VB.OptionButton optActive 
         Caption         =   "Active"
         Height          =   225
         Left            =   240
         TabIndex        =   105
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.TextBox txt_company_name1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   8340
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   101
      Top             =   90
      Width           =   2715
   End
   Begin prj_fej_jkt.vbButton btnChange 
      Height          =   615
      Left            =   5790
      TabIndex        =   100
      Top             =   60
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Change Branch Office"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frm_mst_employee.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   2790
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   88
      Top             =   120
      Width           =   2715
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   54
      Top             =   7800
      Width           =   14175
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Report"
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
         Left            =   10020
         Picture         =   "frm_mst_employee.frx":0028
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_update 
         Caption         =   "&Update"
         Height          =   645
         Left            =   8970
         Picture         =   "frm_mst_employee.frx":05B2
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   3240
         Picture         =   "frm_mst_employee.frx":0B3C
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   4320
         Picture         =   "frm_mst_employee.frx":10C6
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   1080
         Picture         =   "frm_mst_employee.frx":1650
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   13080
         Picture         =   "frm_mst_employee.frx":1BDA
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5400
         Picture         =   "frm_mst_employee.frx":2164
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   2160
         Picture         =   "frm_mst_employee.frx":26EE
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_import 
         Caption         =   "&Import"
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
         Left            =   7920
         Picture         =   "frm_mst_employee.frx":2C78
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   360
         Width           =   975
      End
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Report"
         Height          =   645
         Left            =   0
         Picture         =   "frm_mst_employee.frx":3202
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   2820
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6960
      Top             =   480
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
      Caption         =   "Adodc1"
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1560
      OleObjectBlob   =   "frm_mst_employee.frx":378C
      TabIndex        =   1
      Top             =   120
      Width           =   1155
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company1 
      Height          =   375
      Left            =   7110
      OleObjectBlob   =   "frm_mst_employee.frx":56F2
      TabIndex        =   102
      Top             =   90
      Width           =   1155
   End
   Begin VB.Frame fra_entry 
      Height          =   6855
      Left            =   225
      TabIndex        =   0
      Top             =   750
      Width           =   14175
      Begin VB.TextBox txt_npwp_address 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   6120
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox txt_group_shift_name 
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
         Height          =   435
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   115
         Top             =   3600
         Width           =   2655
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "MANAGERIAL ACCESS"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   420
         TabIndex        =   113
         Top             =   4410
         Width           =   2265
      End
      Begin VB.Frame frmLoading 
         Height          =   555
         Left            =   5550
         TabIndex        =   111
         Top             =   2940
         Visible         =   0   'False
         Width           =   3735
         Begin VB.Label Label9 
            Caption         =   "Please wait... Update Data Employee !"
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
            Height          =   255
            Left            =   210
            TabIndex        =   112
            Top             =   180
            Width           =   3465
         End
      End
      Begin VB.TextBox txt_branch 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   109
         Top             =   4920
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "MANAGERIAL LEVEL"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   420
         TabIndex        =   7
         Top             =   4110
         Width           =   2265
      End
      Begin VB.Frame Frame1 
         Caption         =   "PARENT"
         Height          =   1815
         Left            =   9240
         TabIndex        =   95
         Top             =   2640
         Width           =   4575
         Begin VB.TextBox txt_child_number_from 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3480
            MaxLength       =   50
            TabIndex        =   37
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txt_child_number 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   36
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txt_mothers_name 
            Appearance      =   0  'Flat
            CausesValidation=   0   'False
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   35
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox txt_fathers_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   34
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "FROM"
            Height          =   195
            Left            =   2760
            TabIndex        =   99
            Top             =   1200
            Width           =   465
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "MOTHER"
            Height          =   195
            Left            =   360
            TabIndex        =   98
            Top             =   840
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "CHILD NO."
            Height          =   195
            Left            =   360
            TabIndex        =   97
            Top             =   1200
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "FATHER"
            Height          =   195
            Left            =   360
            TabIndex        =   96
            Top             =   480
            Width           =   645
         End
      End
      Begin VB.TextBox txt_email 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         TabIndex        =   19
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox txt_number_of_children 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   30
         TabIndex        =   17
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txt_npwp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   30
         TabIndex        =   20
         Top             =   3360
         Width           =   2775
      End
      Begin VB.TextBox txt_employee_nick_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox cbo_shiftable 
         Height          =   315
         ItemData        =   "frm_mst_employee.frx":7659
         Left            =   10800
         List            =   "frm_mst_employee.frx":7663
         TabIndex        =   33
         Text            =   "..."
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Frame fra_last_education 
         Caption         =   "LAST EDUCATION"
         Height          =   1485
         Left            =   4560
         TabIndex        =   84
         Top             =   5340
         Width           =   4335
         Begin VB.TextBox txt_last_education_code_other 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2400
            MaxLength       =   50
            TabIndex        =   25
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox cbo_last_education_code 
            Height          =   315
            ItemData        =   "frm_mst_employee.frx":7670
            Left            =   960
            List            =   "frm_mst_employee.frx":7695
            TabIndex        =   24
            Text            =   "..."
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txt_last_education_name 
            Appearance      =   0  'Flat
            CausesValidation=   0   'False
            Height          =   315
            Left            =   960
            MaxLength       =   50
            TabIndex        =   26
            Top             =   720
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker DTPicker_last_education_pass 
            Height          =   315
            Left            =   960
            TabIndex        =   27
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM"
            Format          =   94109699
            CurrentDate     =   39270
         End
         Begin VB.Label lbl_last_education_level 
            AutoSize        =   -1  'True
            Caption         =   "LEVEL"
            Height          =   195
            Left            =   360
            TabIndex        =   87
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lbl_last_education_pass 
            AutoSize        =   -1  'True
            Caption         =   "PASS"
            Height          =   195
            Left            =   360
            TabIndex        =   86
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label lbl_last_education_name 
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            Height          =   195
            Left            =   360
            TabIndex        =   85
            Top             =   720
            Width           =   465
         End
      End
      Begin VB.ComboBox cbo_religion 
         Height          =   315
         ItemData        =   "frm_mst_employee.frx":76CA
         Left            =   6120
         List            =   "frm_mst_employee.frx":76E0
         TabIndex        =   15
         Text            =   "..."
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Frame fra_last_employment 
         Caption         =   "LAST EMPLOYMENT"
         Height          =   1635
         Left            =   360
         TabIndex        =   79
         Top             =   5130
         Width           =   3855
         Begin VB.TextBox txt_last_employment_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   960
            MaxLength       =   50
            TabIndex        =   8
            Top             =   420
            Width           =   2535
         End
         Begin VB.TextBox txt_last_employment_title 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   960
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1140
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker DTPicker_last_employment 
            Height          =   315
            Left            =   960
            TabIndex        =   9
            Top             =   780
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   94109699
            CurrentDate     =   39270
         End
         Begin VB.Label lbl_last_employment_title 
            AutoSize        =   -1  'True
            Caption         =   "TITLE"
            Height          =   195
            Left            =   360
            TabIndex        =   82
            Top             =   1140
            Width           =   450
         End
         Begin VB.Label lbl_last_employment_date 
            AutoSize        =   -1  'True
            Caption         =   "DATE"
            Height          =   195
            Left            =   360
            TabIndex        =   81
            Top             =   780
            Width           =   435
         End
         Begin VB.Label lbl_last_employment_name 
            AutoSize        =   -1  'True
            Caption         =   "NAME"
            Height          =   195
            Left            =   360
            TabIndex        =   80
            Top             =   420
            Width           =   465
         End
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   10800
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   1560
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
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Frame fra_not_active 
         Height          =   1485
         Left            =   9240
         TabIndex        =   68
         Top             =   5340
         Visible         =   0   'False
         Width           =   4575
         Begin VB.TextBox txt_end_working_reason 
            Appearance      =   0  'Flat
            Height          =   675
            Left            =   1560
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Top             =   720
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker DTPicker_end_working 
            Height          =   315
            Left            =   1560
            TabIndex        =   69
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   94109699
            CurrentDate     =   39270
         End
         Begin VB.Label lbl_reason_description 
            AutoSize        =   -1  'True
            Caption         =   "REASON"
            Height          =   195
            Left            =   240
            TabIndex        =   71
            Top             =   720
            Width           =   675
         End
         Begin VB.Label lbl_reason_end_working 
            AutoSize        =   -1  'True
            Caption         =   "END WORKING"
            Height          =   195
            Left            =   240
            TabIndex        =   70
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label lbl_end_mc 
            Caption         =   "END MC"
            Height          =   315
            Left            =   240
            TabIndex        =   114
            Top             =   390
            Width           =   945
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Working Status Employee"
         Height          =   735
         Left            =   9240
         TabIndex        =   73
         Top             =   4680
         Width           =   4575
         Begin VB.OptionButton opt_mc 
            Caption         =   "Medical Check"
            Height          =   255
            Left            =   2940
            TabIndex        =   103
            Top             =   330
            Width           =   1545
         End
         Begin VB.OptionButton opt_not_active 
            Caption         =   "NOT ACTIVE"
            Height          =   255
            Left            =   1440
            TabIndex        =   75
            Top             =   330
            Width           =   1575
         End
         Begin VB.OptionButton opt_active 
            Caption         =   "ACTIVE"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   330
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txt_working_time 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   13320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   29
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txt_title_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   13320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   66
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txt_bank_account 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   30
         TabIndex        =   23
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox txt_phone_number 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   30
         TabIndex        =   22
         Top             =   4200
         Width           =   2775
      End
      Begin VB.ComboBox cbo_marital_status 
         Height          =   315
         ItemData        =   "frm_mst_employee.frx":7714
         Left            =   6120
         List            =   "frm_mst_employee.frx":7724
         TabIndex        =   16
         Text            =   "..."
         Top             =   1800
         Width           =   1695
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
         Height          =   435
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   59
         Top             =   2760
         Width           =   2655
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
         Height          =   435
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txt_employee_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         MaxLength       =   30
         TabIndex        =   2
         Top             =   480
         Width           =   2655
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
         TabIndex        =   48
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_place_of_birth 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   30
         TabIndex        =   13
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txt_address 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   6120
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         MaxLength       =   100
         TabIndex        =   3
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cbo_sex 
         Height          =   315
         ItemData        =   "frm_mst_employee.frx":7749
         Left            =   6120
         List            =   "frm_mst_employee.frx":7753
         TabIndex        =   14
         Text            =   "..."
         Top             =   1080
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker_birth 
         Height          =   315
         Left            =   6120
         TabIndex        =   11
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   94109699
         CurrentDate     =   39270
      End
      Begin MSComCtl2.DTPicker DTPicker_start_working 
         Height          =   315
         Left            =   10800
         TabIndex        =   28
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   94109699
         CurrentDate     =   39270
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_title 
         Height          =   375
         Left            =   10800
         OleObjectBlob   =   "frm_mst_employee.frx":7765
         TabIndex        =   31
         Top             =   1200
         Width           =   2475
      End
      Begin MSComCtl2.DTPicker DTPicker_appointment 
         Height          =   315
         Left            =   10800
         TabIndex        =   30
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   94109699
         CurrentDate     =   39270
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_department 
         Height          =   375
         Left            =   1590
         OleObjectBlob   =   "frm_mst_employee.frx":96BD
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_division 
         Height          =   375
         Left            =   1590
         OleObjectBlob   =   "frm_mst_employee.frx":B626
         TabIndex        =   6
         Top             =   2400
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Adodc_department 
         Height          =   375
         Left            =   2070
         Top             =   1620
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
      Begin MSAdodcLib.Adodc Adodc_division 
         Height          =   375
         Left            =   2070
         Top             =   2400
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
      Begin MSAdodcLib.Adodc Adodc_title 
         Height          =   330
         Left            =   9960
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
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
      Begin TrueOleDBList60.TDBCombo TDBCombo_group_shift 
         Height          =   375
         Left            =   1590
         OleObjectBlob   =   "frm_mst_employee.frx":D58D
         TabIndex        =   116
         Top             =   3240
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Adodc_group_shift 
         Height          =   375
         Left            =   2070
         Top             =   3240
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
      Begin MSComCtl2.DTPicker dt_end_mc 
         Height          =   315
         Left            =   10800
         TabIndex        =   38
         Top             =   5700
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94109697
         CurrentDate     =   40849
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "END MC"
         Height          =   195
         Left            =   9810
         TabIndex        =   119
         Top             =   5730
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "NPWP ADDRESS"
         Height          =   195
         Left            =   4560
         TabIndex        =   118
         Top             =   3750
         Width           =   1320
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "GROUP SHIFT*"
         Height          =   195
         Left            =   390
         TabIndex        =   117
         Top             =   3240
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "BRANCHES"
         Height          =   195
         Left            =   4560
         TabIndex        =   110
         Top             =   4920
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "EMAIL"
         Height          =   195
         Left            =   4560
         TabIndex        =   94
         Top             =   3000
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CHILDREN'S*"
         Height          =   195
         Left            =   4560
         TabIndex        =   93
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N P W P*"
         Height          =   195
         Left            =   4560
         TabIndex        =   92
         Top             =   3360
         Width           =   690
      End
      Begin VB.Label lbl_employee_nick_name 
         AutoSize        =   -1  'True
         Caption         =   "NICK NAME*"
         Height          =   195
         Left            =   390
         TabIndex        =   91
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label lbl_shiftable 
         AutoSize        =   -1  'True
         Caption         =   "SHIFTABLE"
         Height          =   195
         Left            =   9240
         TabIndex        =   90
         Top             =   2160
         Width           =   870
      End
      Begin VB.Label lbl_religion 
         AutoSize        =   -1  'True
         Caption         =   "RELIGION*"
         Height          =   195
         Left            =   4560
         TabIndex        =   83
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label lbl_description 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
         Height          =   195
         Left            =   9240
         TabIndex        =   78
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lbl_appointment 
         AutoSize        =   -1  'True
         Caption         =   "APPOINTMENT"
         Height          =   195
         Left            =   9240
         TabIndex        =   77
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "AGE (Y)*"
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   7680
         TabIndex        =   76
         Top             =   420
         Width           =   630
      End
      Begin VB.Label lbl_working_age 
         AutoSize        =   -1  'True
         Caption         =   "W. AGE (Y)*"
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   12360
         TabIndex        =   72
         Top             =   480
         Width           =   885
      End
      Begin VB.Label lbl_title 
         AutoSize        =   -1  'True
         Caption         =   "TITLE*"
         Height          =   195
         Left            =   9240
         TabIndex        =   67
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label lbl_start_working 
         AutoSize        =   -1  'True
         Caption         =   "START WORKING*"
         Height          =   195
         Left            =   9240
         TabIndex        =   65
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label lbl_bank_account 
         AutoSize        =   -1  'True
         Caption         =   "BANK ACCOUNT"
         Height          =   195
         Left            =   4560
         TabIndex        =   64
         Top             =   4560
         Width           =   1260
      End
      Begin VB.Label lbl_phone_number 
         AutoSize        =   -1  'True
         Caption         =   "PHONE NUMBER"
         Height          =   195
         Left            =   4560
         TabIndex        =   63
         Top             =   4200
         Width           =   1320
      End
      Begin VB.Label lbl_marital_status 
         AutoSize        =   -1  'True
         Caption         =   "MARITAL STATUS*"
         Height          =   195
         Left            =   4560
         TabIndex        =   62
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "SEX*"
         Height          =   195
         Left            =   4560
         TabIndex        =   61
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lbl_division 
         AutoSize        =   -1  'True
         Caption         =   "DIVISION*"
         Height          =   195
         Left            =   390
         TabIndex        =   60
         Top             =   2400
         Width           =   765
      End
      Begin VB.Label lbl_department 
         AutoSize        =   -1  'True
         Caption         =   "DEPARTMENT*"
         Height          =   195
         Left            =   390
         TabIndex        =   58
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label lbl_employee_name 
         AutoSize        =   -1  'True
         Caption         =   "EMP. NAME*"
         Height          =   195
         Left            =   390
         TabIndex        =   56
         Top             =   840
         Width           =   960
      End
      Begin VB.Label lbl_address 
         AutoSize        =   -1  'True
         Caption         =   "ADDRESS"
         Height          =   195
         Left            =   4560
         TabIndex        =   52
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label lbl_employee_code 
         AutoSize        =   -1  'True
         Caption         =   "NO. BADGE*"
         Height          =   195
         Left            =   390
         TabIndex        =   51
         Top             =   480
         Width           =   945
      End
      Begin VB.Label lbl_date_of_birth 
         AutoSize        =   -1  'True
         Caption         =   "DATE OF BIRTH"
         Height          =   195
         Left            =   4560
         TabIndex        =   50
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label lbl_place_of_birth 
         AutoSize        =   -1  'True
         Caption         =   "PLACE OF BIRTH"
         Height          =   195
         Left            =   4560
         TabIndex        =   49
         Top             =   720
         Width           =   1305
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6615
      Left            =   240
      TabIndex        =   53
      Top             =   840
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11668
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
      Columns(2).Caption=   "DEPT. CODE"
      Columns(2).DataField=   "department_code"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "DEPT. NAME"
      Columns(3).DataField=   "department_name"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "DIV. CODE"
      Columns(4).DataField=   "division_code"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DIV. NAME"
      Columns(5).DataField=   "division_name"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "EMP. CODE"
      Columns(6).DataField=   "employee_code"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "EMP. NAME"
      Columns(7).DataField=   "employee_name"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "NICK NAME"
      Columns(8).DataField=   "employee_nick_name"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "GROUP SHIFT"
      Columns(9).DataField=   "group_name"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   4
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "ACTIVE"
      Columns(10).DataField=   "flag_active"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "EMAIL"
      Columns(11).DataField=   "email"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "BIRTH DATE"
      Columns(12).DataField=   "date_of_birth"
      Columns(12).NumberFormat=   "FormatText Event"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "PLACE OF BIRTH"
      Columns(13).DataField=   "place_of_birth"
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
      Columns(14).Caption=   "SEX"
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
      Columns(16).Caption=   "ADDRESS"
      Columns(16).DataField=   "address"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "PHONE NUMBER"
      Columns(17).DataField=   "phone_number"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "NPWP"
      Columns(18).DataField=   "npwp"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "BANK ACCOUNT"
      Columns(19).DataField=   "bank_account"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "START WORKING"
      Columns(20).DataField=   "start_working"
      Columns(20).NumberFormat=   "FormatText Event"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "TITLE CODE"
      Columns(21).DataField=   "title_code"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "TITLE NAME"
      Columns(22).DataField=   "title_name"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   4
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "SHIFTABLE"
      Columns(23).DataField=   "flag_shiftable"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "END WORKING"
      Columns(24).DataField=   "end_working"
      Columns(24).NumberFormat=   "FormatText Event"
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "REASON"
      Columns(25).DataField=   "reason"
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   26
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=26"
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
      Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(30)=   "Column(4).Width=1508"
      Splits(0)._ColumnProps(31)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._WidthInPix=1429"
      Splits(0)._ColumnProps(33)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(35)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(36)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(38)=   "Column(5).Width=2514"
      Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=2434"
      Splits(0)._ColumnProps(41)=   "Column(5)._ColStyle=8708"
      Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(43)=   "Column(6).Width=1588"
      Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=1508"
      Splits(0)._ColumnProps(46)=   "Column(6).AllowSizing=0"
      Splits(0)._ColumnProps(47)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(48)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(49)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(50)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(51)=   "Column(7).Width=1588"
      Splits(0)._ColumnProps(52)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(7)._WidthInPix=1508"
      Splits(0)._ColumnProps(54)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(55)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(56)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(57)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(58)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(59)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(60)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(62)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(63)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(64)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(65)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(66)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(67)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(68)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(70)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(71)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(72)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(73)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(75)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(76)=   "Column(10)._ColStyle=513"
      Splits(0)._ColumnProps(77)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(78)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(79)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(80)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(81)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(82)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(83)=   "Column(11).AllowSizing=0"
      Splits(0)._ColumnProps(84)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(85)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(86)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(87)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(88)=   "Column(12).Width=2064"
      Splits(0)._ColumnProps(89)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(90)=   "Column(12)._WidthInPix=1984"
      Splits(0)._ColumnProps(91)=   "Column(12).AllowSizing=0"
      Splits(0)._ColumnProps(92)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(93)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(94)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(95)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(96)=   "Column(13).Width=3016"
      Splits(0)._ColumnProps(97)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(98)=   "Column(13)._WidthInPix=2937"
      Splits(0)._ColumnProps(99)=   "Column(13).AllowSizing=0"
      Splits(0)._ColumnProps(100)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(101)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(102)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(103)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(104)=   "Column(14).Width=2037"
      Splits(0)._ColumnProps(105)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(106)=   "Column(14)._WidthInPix=1958"
      Splits(0)._ColumnProps(107)=   "Column(14).AllowSizing=0"
      Splits(0)._ColumnProps(108)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(109)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(110)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(111)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(112)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(113)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(114)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(115)=   "Column(15).AllowSizing=0"
      Splits(0)._ColumnProps(116)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(117)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(118)=   "Column(15).AllowFocus=0"
      Splits(0)._ColumnProps(119)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(120)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(121)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(122)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(123)=   "Column(16).AllowSizing=0"
      Splits(0)._ColumnProps(124)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(125)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(126)=   "Column(16).AllowFocus=0"
      Splits(0)._ColumnProps(127)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(128)=   "Column(16)._MinWidth=10"
      Splits(0)._ColumnProps(129)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(130)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(131)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(132)=   "Column(17).AllowSizing=0"
      Splits(0)._ColumnProps(133)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(134)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(135)=   "Column(17).AllowFocus=0"
      Splits(0)._ColumnProps(136)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(137)=   "Column(17)._MinWidth=54215968"
      Splits(0)._ColumnProps(138)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(139)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(140)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(141)=   "Column(18)._ColStyle=516"
      Splits(0)._ColumnProps(142)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(143)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(144)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(145)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(146)=   "Column(19).AllowSizing=0"
      Splits(0)._ColumnProps(147)=   "Column(19)._ColStyle=516"
      Splits(0)._ColumnProps(148)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(149)=   "Column(19).AllowFocus=0"
      Splits(0)._ColumnProps(150)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(151)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(152)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(153)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(154)=   "Column(20).AllowSizing=0"
      Splits(0)._ColumnProps(155)=   "Column(20)._ColStyle=516"
      Splits(0)._ColumnProps(156)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(157)=   "Column(20).AllowFocus=0"
      Splits(0)._ColumnProps(158)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(159)=   "Column(20)._MinWidth=60129312"
      Splits(0)._ColumnProps(160)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(161)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(162)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(163)=   "Column(21).AllowSizing=0"
      Splits(0)._ColumnProps(164)=   "Column(21)._ColStyle=516"
      Splits(0)._ColumnProps(165)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(166)=   "Column(21).AllowFocus=0"
      Splits(0)._ColumnProps(167)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(168)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(169)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(170)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(171)=   "Column(22).AllowSizing=0"
      Splits(0)._ColumnProps(172)=   "Column(22)._ColStyle=516"
      Splits(0)._ColumnProps(173)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(174)=   "Column(22).AllowFocus=0"
      Splits(0)._ColumnProps(175)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(176)=   "Column(22)._MinWidth=79702332"
      Splits(0)._ColumnProps(177)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(178)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(179)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(180)=   "Column(23).AllowSizing=0"
      Splits(0)._ColumnProps(181)=   "Column(23)._ColStyle=516"
      Splits(0)._ColumnProps(182)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(183)=   "Column(23).AllowFocus=0"
      Splits(0)._ColumnProps(184)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(185)=   "Column(23)._MinWidth=79897920"
      Splits(0)._ColumnProps(186)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(187)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(188)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(189)=   "Column(24).AllowSizing=0"
      Splits(0)._ColumnProps(190)=   "Column(24)._ColStyle=516"
      Splits(0)._ColumnProps(191)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(192)=   "Column(24).AllowFocus=0"
      Splits(0)._ColumnProps(193)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(194)=   "Column(24)._MinWidth=79914544"
      Splits(0)._ColumnProps(195)=   "Column(25).Width=2725"
      Splits(0)._ColumnProps(196)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(197)=   "Column(25)._WidthInPix=2646"
      Splits(0)._ColumnProps(198)=   "Column(25).AllowSizing=0"
      Splits(0)._ColumnProps(199)=   "Column(25)._ColStyle=516"
      Splits(0)._ColumnProps(200)=   "Column(25).Visible=0"
      Splits(0)._ColumnProps(201)=   "Column(25).AllowFocus=0"
      Splits(0)._ColumnProps(202)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(203)=   "Column(25)._MinWidth=79789632"
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
      Splits(1)._ColumnProps(0)=   "Columns.Count=26"
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
      Splits(1)._ColumnProps(51)=   "Column(6).Width=1879"
      Splits(1)._ColumnProps(52)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(53)=   "Column(6)._WidthInPix=1799"
      Splits(1)._ColumnProps(54)=   "Column(6)._ColStyle=8708"
      Splits(1)._ColumnProps(55)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(56)=   "Column(6)._MinWidth=80000960"
      Splits(1)._ColumnProps(57)=   "Column(7).Width=3678"
      Splits(1)._ColumnProps(58)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(59)=   "Column(7)._WidthInPix=3598"
      Splits(1)._ColumnProps(60)=   "Column(7)._ColStyle=8708"
      Splits(1)._ColumnProps(61)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(62)=   "Column(7)._MinWidth=79999936"
      Splits(1)._ColumnProps(63)=   "Column(8).Width=2170"
      Splits(1)._ColumnProps(64)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(65)=   "Column(8)._WidthInPix=2090"
      Splits(1)._ColumnProps(66)=   "Column(8)._ColStyle=8708"
      Splits(1)._ColumnProps(67)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(68)=   "Column(8)._MinWidth=80007280"
      Splits(1)._ColumnProps(69)=   "Column(9).Width=2725"
      Splits(1)._ColumnProps(70)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(71)=   "Column(9)._WidthInPix=2646"
      Splits(1)._ColumnProps(72)=   "Column(9)._ColStyle=516"
      Splits(1)._ColumnProps(73)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(74)=   "Column(9)._MinWidth=80007280"
      Splits(1)._ColumnProps(75)=   "Column(10).Width=1191"
      Splits(1)._ColumnProps(76)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(77)=   "Column(10)._WidthInPix=1111"
      Splits(1)._ColumnProps(78)=   "Column(10)._ColStyle=513"
      Splits(1)._ColumnProps(79)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(80)=   "Column(10)._MinWidth=80007280"
      Splits(1)._ColumnProps(81)=   "Column(11).Width=4233"
      Splits(1)._ColumnProps(82)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(83)=   "Column(11)._WidthInPix=4154"
      Splits(1)._ColumnProps(84)=   "Column(11)._ColStyle=516"
      Splits(1)._ColumnProps(85)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(86)=   "Column(11)._MinWidth=80007280"
      Splits(1)._ColumnProps(87)=   "Column(12).Width=2064"
      Splits(1)._ColumnProps(88)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(89)=   "Column(12)._WidthInPix=1984"
      Splits(1)._ColumnProps(90)=   "Column(12)._ColStyle=8705"
      Splits(1)._ColumnProps(91)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(92)=   "Column(12)._MinWidth=80007280"
      Splits(1)._ColumnProps(93)=   "Column(13).Width=3016"
      Splits(1)._ColumnProps(94)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(95)=   "Column(13)._WidthInPix=2937"
      Splits(1)._ColumnProps(96)=   "Column(13)._ColStyle=8708"
      Splits(1)._ColumnProps(97)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(98)=   "Column(13)._MinWidth=80010048"
      Splits(1)._ColumnProps(99)=   "Column(14).Width=2037"
      Splits(1)._ColumnProps(100)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(101)=   "Column(14)._WidthInPix=1958"
      Splits(1)._ColumnProps(102)=   "Column(14)._ColStyle=8705"
      Splits(1)._ColumnProps(103)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(104)=   "Column(15).Width=2725"
      Splits(1)._ColumnProps(105)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(106)=   "Column(15)._WidthInPix=2646"
      Splits(1)._ColumnProps(107)=   "Column(15)._ColStyle=8705"
      Splits(1)._ColumnProps(108)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(109)=   "Column(16).Width=2725"
      Splits(1)._ColumnProps(110)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(111)=   "Column(16)._WidthInPix=2646"
      Splits(1)._ColumnProps(112)=   "Column(16)._ColStyle=8708"
      Splits(1)._ColumnProps(113)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(114)=   "Column(17).Width=2725"
      Splits(1)._ColumnProps(115)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(116)=   "Column(17)._WidthInPix=2646"
      Splits(1)._ColumnProps(117)=   "Column(17)._ColStyle=8708"
      Splits(1)._ColumnProps(118)=   "Column(17).Order=18"
      Splits(1)._ColumnProps(119)=   "Column(18).Width=2725"
      Splits(1)._ColumnProps(120)=   "Column(18).DividerColor=0"
      Splits(1)._ColumnProps(121)=   "Column(18)._WidthInPix=2646"
      Splits(1)._ColumnProps(122)=   "Column(18)._ColStyle=516"
      Splits(1)._ColumnProps(123)=   "Column(18).Order=19"
      Splits(1)._ColumnProps(124)=   "Column(19).Width=2725"
      Splits(1)._ColumnProps(125)=   "Column(19).DividerColor=0"
      Splits(1)._ColumnProps(126)=   "Column(19)._WidthInPix=2646"
      Splits(1)._ColumnProps(127)=   "Column(19)._ColStyle=8708"
      Splits(1)._ColumnProps(128)=   "Column(19).Order=20"
      Splits(1)._ColumnProps(129)=   "Column(20).Width=2725"
      Splits(1)._ColumnProps(130)=   "Column(20).DividerColor=0"
      Splits(1)._ColumnProps(131)=   "Column(20)._WidthInPix=2646"
      Splits(1)._ColumnProps(132)=   "Column(20)._ColStyle=8705"
      Splits(1)._ColumnProps(133)=   "Column(20).Order=21"
      Splits(1)._ColumnProps(134)=   "Column(21).Width=2725"
      Splits(1)._ColumnProps(135)=   "Column(21).DividerColor=0"
      Splits(1)._ColumnProps(136)=   "Column(21)._WidthInPix=2646"
      Splits(1)._ColumnProps(137)=   "Column(21)._ColStyle=8708"
      Splits(1)._ColumnProps(138)=   "Column(21).Order=22"
      Splits(1)._ColumnProps(139)=   "Column(22).Width=2725"
      Splits(1)._ColumnProps(140)=   "Column(22).DividerColor=0"
      Splits(1)._ColumnProps(141)=   "Column(22)._WidthInPix=2646"
      Splits(1)._ColumnProps(142)=   "Column(22)._ColStyle=8708"
      Splits(1)._ColumnProps(143)=   "Column(22).Order=23"
      Splits(1)._ColumnProps(144)=   "Column(23).Width=1720"
      Splits(1)._ColumnProps(145)=   "Column(23).DividerColor=0"
      Splits(1)._ColumnProps(146)=   "Column(23)._WidthInPix=1640"
      Splits(1)._ColumnProps(147)=   "Column(23)._ColStyle=8705"
      Splits(1)._ColumnProps(148)=   "Column(23).Order=24"
      Splits(1)._ColumnProps(149)=   "Column(24).Width=2725"
      Splits(1)._ColumnProps(150)=   "Column(24).DividerColor=0"
      Splits(1)._ColumnProps(151)=   "Column(24)._WidthInPix=2646"
      Splits(1)._ColumnProps(152)=   "Column(24)._ColStyle=8705"
      Splits(1)._ColumnProps(153)=   "Column(24).Order=25"
      Splits(1)._ColumnProps(154)=   "Column(25).Width=2725"
      Splits(1)._ColumnProps(155)=   "Column(25).DividerColor=0"
      Splits(1)._ColumnProps(156)=   "Column(25)._WidthInPix=2646"
      Splits(1)._ColumnProps(157)=   "Column(25)._ColStyle=8708"
      Splits(1)._ColumnProps(158)=   "Column(25).Order=26"
      Splits(1)._ColumnProps(159)=   "Column(25)._MinWidth=80015760"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   3
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
      Caption         =   "LIST OF EMPLOYEE"
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
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=98,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=95,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=96,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=97,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=250,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=247,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=248,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=249,.parent=17"
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
      _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=222,.parent=13"
      _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=219,.parent=14"
      _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=220,.parent=15"
      _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=221,.parent=17"
      _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=46,.parent=13"
      _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=43,.parent=14"
      _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=44,.parent=15"
      _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=45,.parent=17"
      _StyleDefs(114) =   "Splits(0).Columns(20).Style:id=58,.parent=13"
      _StyleDefs(115) =   "Splits(0).Columns(20).HeadingStyle:id=55,.parent=14"
      _StyleDefs(116) =   "Splits(0).Columns(20).FooterStyle:id=56,.parent=15"
      _StyleDefs(117) =   "Splits(0).Columns(20).EditorStyle:id=57,.parent=17"
      _StyleDefs(118) =   "Splits(0).Columns(21).Style:id=94,.parent=13"
      _StyleDefs(119) =   "Splits(0).Columns(21).HeadingStyle:id=91,.parent=14"
      _StyleDefs(120) =   "Splits(0).Columns(21).FooterStyle:id=92,.parent=15"
      _StyleDefs(121) =   "Splits(0).Columns(21).EditorStyle:id=93,.parent=17"
      _StyleDefs(122) =   "Splits(0).Columns(22).Style:id=106,.parent=13"
      _StyleDefs(123) =   "Splits(0).Columns(22).HeadingStyle:id=103,.parent=14"
      _StyleDefs(124) =   "Splits(0).Columns(22).FooterStyle:id=104,.parent=15"
      _StyleDefs(125) =   "Splits(0).Columns(22).EditorStyle:id=105,.parent=17"
      _StyleDefs(126) =   "Splits(0).Columns(23).Style:id=114,.parent=13"
      _StyleDefs(127) =   "Splits(0).Columns(23).HeadingStyle:id=111,.parent=14"
      _StyleDefs(128) =   "Splits(0).Columns(23).FooterStyle:id=112,.parent=15"
      _StyleDefs(129) =   "Splits(0).Columns(23).EditorStyle:id=113,.parent=17"
      _StyleDefs(130) =   "Splits(0).Columns(24).Style:id=118,.parent=13"
      _StyleDefs(131) =   "Splits(0).Columns(24).HeadingStyle:id=115,.parent=14"
      _StyleDefs(132) =   "Splits(0).Columns(24).FooterStyle:id=116,.parent=15"
      _StyleDefs(133) =   "Splits(0).Columns(24).EditorStyle:id=117,.parent=17"
      _StyleDefs(134) =   "Splits(0).Columns(25).Style:id=122,.parent=13"
      _StyleDefs(135) =   "Splits(0).Columns(25).HeadingStyle:id=119,.parent=14"
      _StyleDefs(136) =   "Splits(0).Columns(25).FooterStyle:id=120,.parent=15"
      _StyleDefs(137) =   "Splits(0).Columns(25).EditorStyle:id=121,.parent=17"
      _StyleDefs(138) =   "Splits(1).Style:id=123,.parent=1"
      _StyleDefs(139) =   "Splits(1).CaptionStyle:id=132,.parent=4,.bgcolor=&H80000002&"
      _StyleDefs(140) =   ":id=132,.fgcolor=&H80000009&"
      _StyleDefs(141) =   "Splits(1).HeadingStyle:id=124,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(142) =   ":id=124,.fgcolor=&H80000002&"
      _StyleDefs(143) =   "Splits(1).FooterStyle:id=125,.parent=3"
      _StyleDefs(144) =   "Splits(1).InactiveStyle:id=126,.parent=5"
      _StyleDefs(145) =   "Splits(1).SelectedStyle:id=128,.parent=6"
      _StyleDefs(146) =   "Splits(1).EditorStyle:id=127,.parent=7"
      _StyleDefs(147) =   "Splits(1).HighlightRowStyle:id=129,.parent=8"
      _StyleDefs(148) =   "Splits(1).EvenRowStyle:id=130,.parent=9"
      _StyleDefs(149) =   "Splits(1).OddRowStyle:id=131,.parent=10"
      _StyleDefs(150) =   "Splits(1).RecordSelectorStyle:id=133,.parent=11"
      _StyleDefs(151) =   "Splits(1).FilterBarStyle:id=134,.parent=12"
      _StyleDefs(152) =   "Splits(1).Columns(0).Style:id=138,.parent=123"
      _StyleDefs(153) =   "Splits(1).Columns(0).HeadingStyle:id=135,.parent=124"
      _StyleDefs(154) =   "Splits(1).Columns(0).FooterStyle:id=136,.parent=125"
      _StyleDefs(155) =   "Splits(1).Columns(0).EditorStyle:id=137,.parent=127"
      _StyleDefs(156) =   "Splits(1).Columns(1).Style:id=142,.parent=123"
      _StyleDefs(157) =   "Splits(1).Columns(1).HeadingStyle:id=139,.parent=124"
      _StyleDefs(158) =   "Splits(1).Columns(1).FooterStyle:id=140,.parent=125"
      _StyleDefs(159) =   "Splits(1).Columns(1).EditorStyle:id=141,.parent=127"
      _StyleDefs(160) =   "Splits(1).Columns(2).Style:id=146,.parent=123"
      _StyleDefs(161) =   "Splits(1).Columns(2).HeadingStyle:id=143,.parent=124"
      _StyleDefs(162) =   "Splits(1).Columns(2).FooterStyle:id=144,.parent=125"
      _StyleDefs(163) =   "Splits(1).Columns(2).EditorStyle:id=145,.parent=127"
      _StyleDefs(164) =   "Splits(1).Columns(3).Style:id=150,.parent=123"
      _StyleDefs(165) =   "Splits(1).Columns(3).HeadingStyle:id=147,.parent=124"
      _StyleDefs(166) =   "Splits(1).Columns(3).FooterStyle:id=148,.parent=125"
      _StyleDefs(167) =   "Splits(1).Columns(3).EditorStyle:id=149,.parent=127"
      _StyleDefs(168) =   "Splits(1).Columns(4).Style:id=154,.parent=123"
      _StyleDefs(169) =   "Splits(1).Columns(4).HeadingStyle:id=151,.parent=124"
      _StyleDefs(170) =   "Splits(1).Columns(4).FooterStyle:id=152,.parent=125"
      _StyleDefs(171) =   "Splits(1).Columns(4).EditorStyle:id=153,.parent=127"
      _StyleDefs(172) =   "Splits(1).Columns(5).Style:id=158,.parent=123"
      _StyleDefs(173) =   "Splits(1).Columns(5).HeadingStyle:id=155,.parent=124"
      _StyleDefs(174) =   "Splits(1).Columns(5).FooterStyle:id=156,.parent=125"
      _StyleDefs(175) =   "Splits(1).Columns(5).EditorStyle:id=157,.parent=127"
      _StyleDefs(176) =   "Splits(1).Columns(6).Style:id=162,.parent=123,.locked=-1"
      _StyleDefs(177) =   "Splits(1).Columns(6).HeadingStyle:id=159,.parent=124"
      _StyleDefs(178) =   "Splits(1).Columns(6).FooterStyle:id=160,.parent=125"
      _StyleDefs(179) =   "Splits(1).Columns(6).EditorStyle:id=161,.parent=127"
      _StyleDefs(180) =   "Splits(1).Columns(7).Style:id=166,.parent=123,.locked=-1"
      _StyleDefs(181) =   "Splits(1).Columns(7).HeadingStyle:id=163,.parent=124"
      _StyleDefs(182) =   "Splits(1).Columns(7).FooterStyle:id=164,.parent=125"
      _StyleDefs(183) =   "Splits(1).Columns(7).EditorStyle:id=165,.parent=127"
      _StyleDefs(184) =   "Splits(1).Columns(8).Style:id=230,.parent=123,.locked=-1"
      _StyleDefs(185) =   "Splits(1).Columns(8).HeadingStyle:id=227,.parent=124"
      _StyleDefs(186) =   "Splits(1).Columns(8).FooterStyle:id=228,.parent=125"
      _StyleDefs(187) =   "Splits(1).Columns(8).EditorStyle:id=229,.parent=127"
      _StyleDefs(188) =   "Splits(1).Columns(9).Style:id=254,.parent=123"
      _StyleDefs(189) =   "Splits(1).Columns(9).HeadingStyle:id=251,.parent=124"
      _StyleDefs(190) =   "Splits(1).Columns(9).FooterStyle:id=252,.parent=125"
      _StyleDefs(191) =   "Splits(1).Columns(9).EditorStyle:id=253,.parent=127"
      _StyleDefs(192) =   "Splits(1).Columns(10).Style:id=238,.parent=123,.alignment=2"
      _StyleDefs(193) =   "Splits(1).Columns(10).HeadingStyle:id=235,.parent=124"
      _StyleDefs(194) =   "Splits(1).Columns(10).FooterStyle:id=236,.parent=125"
      _StyleDefs(195) =   "Splits(1).Columns(10).EditorStyle:id=237,.parent=127"
      _StyleDefs(196) =   "Splits(1).Columns(11).Style:id=246,.parent=123"
      _StyleDefs(197) =   "Splits(1).Columns(11).HeadingStyle:id=243,.parent=124"
      _StyleDefs(198) =   "Splits(1).Columns(11).FooterStyle:id=244,.parent=125"
      _StyleDefs(199) =   "Splits(1).Columns(11).EditorStyle:id=245,.parent=127"
      _StyleDefs(200) =   "Splits(1).Columns(12).Style:id=170,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(201) =   "Splits(1).Columns(12).HeadingStyle:id=167,.parent=124"
      _StyleDefs(202) =   "Splits(1).Columns(12).FooterStyle:id=168,.parent=125"
      _StyleDefs(203) =   "Splits(1).Columns(12).EditorStyle:id=169,.parent=127"
      _StyleDefs(204) =   "Splits(1).Columns(13).Style:id=174,.parent=123,.locked=-1"
      _StyleDefs(205) =   "Splits(1).Columns(13).HeadingStyle:id=171,.parent=124"
      _StyleDefs(206) =   "Splits(1).Columns(13).FooterStyle:id=172,.parent=125"
      _StyleDefs(207) =   "Splits(1).Columns(13).EditorStyle:id=173,.parent=127"
      _StyleDefs(208) =   "Splits(1).Columns(14).Style:id=178,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(209) =   "Splits(1).Columns(14).HeadingStyle:id=175,.parent=124"
      _StyleDefs(210) =   "Splits(1).Columns(14).FooterStyle:id=176,.parent=125"
      _StyleDefs(211) =   "Splits(1).Columns(14).EditorStyle:id=177,.parent=127"
      _StyleDefs(212) =   "Splits(1).Columns(15).Style:id=182,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(213) =   "Splits(1).Columns(15).HeadingStyle:id=179,.parent=124"
      _StyleDefs(214) =   "Splits(1).Columns(15).FooterStyle:id=180,.parent=125"
      _StyleDefs(215) =   "Splits(1).Columns(15).EditorStyle:id=181,.parent=127"
      _StyleDefs(216) =   "Splits(1).Columns(16).Style:id=186,.parent=123,.locked=-1"
      _StyleDefs(217) =   "Splits(1).Columns(16).HeadingStyle:id=183,.parent=124"
      _StyleDefs(218) =   "Splits(1).Columns(16).FooterStyle:id=184,.parent=125"
      _StyleDefs(219) =   "Splits(1).Columns(16).EditorStyle:id=185,.parent=127"
      _StyleDefs(220) =   "Splits(1).Columns(17).Style:id=190,.parent=123,.locked=-1"
      _StyleDefs(221) =   "Splits(1).Columns(17).HeadingStyle:id=187,.parent=124"
      _StyleDefs(222) =   "Splits(1).Columns(17).FooterStyle:id=188,.parent=125"
      _StyleDefs(223) =   "Splits(1).Columns(17).EditorStyle:id=189,.parent=127"
      _StyleDefs(224) =   "Splits(1).Columns(18).Style:id=226,.parent=123"
      _StyleDefs(225) =   "Splits(1).Columns(18).HeadingStyle:id=223,.parent=124"
      _StyleDefs(226) =   "Splits(1).Columns(18).FooterStyle:id=224,.parent=125"
      _StyleDefs(227) =   "Splits(1).Columns(18).EditorStyle:id=225,.parent=127"
      _StyleDefs(228) =   "Splits(1).Columns(19).Style:id=194,.parent=123,.locked=-1"
      _StyleDefs(229) =   "Splits(1).Columns(19).HeadingStyle:id=191,.parent=124"
      _StyleDefs(230) =   "Splits(1).Columns(19).FooterStyle:id=192,.parent=125"
      _StyleDefs(231) =   "Splits(1).Columns(19).EditorStyle:id=193,.parent=127"
      _StyleDefs(232) =   "Splits(1).Columns(20).Style:id=198,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(233) =   "Splits(1).Columns(20).HeadingStyle:id=195,.parent=124"
      _StyleDefs(234) =   "Splits(1).Columns(20).FooterStyle:id=196,.parent=125"
      _StyleDefs(235) =   "Splits(1).Columns(20).EditorStyle:id=197,.parent=127"
      _StyleDefs(236) =   "Splits(1).Columns(21).Style:id=202,.parent=123,.locked=-1"
      _StyleDefs(237) =   "Splits(1).Columns(21).HeadingStyle:id=199,.parent=124"
      _StyleDefs(238) =   "Splits(1).Columns(21).FooterStyle:id=200,.parent=125"
      _StyleDefs(239) =   "Splits(1).Columns(21).EditorStyle:id=201,.parent=127"
      _StyleDefs(240) =   "Splits(1).Columns(22).Style:id=206,.parent=123,.locked=-1"
      _StyleDefs(241) =   "Splits(1).Columns(22).HeadingStyle:id=203,.parent=124"
      _StyleDefs(242) =   "Splits(1).Columns(22).FooterStyle:id=204,.parent=125"
      _StyleDefs(243) =   "Splits(1).Columns(22).EditorStyle:id=205,.parent=127"
      _StyleDefs(244) =   "Splits(1).Columns(23).Style:id=210,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(245) =   "Splits(1).Columns(23).HeadingStyle:id=207,.parent=124"
      _StyleDefs(246) =   "Splits(1).Columns(23).FooterStyle:id=208,.parent=125"
      _StyleDefs(247) =   "Splits(1).Columns(23).EditorStyle:id=209,.parent=127"
      _StyleDefs(248) =   "Splits(1).Columns(24).Style:id=214,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(249) =   "Splits(1).Columns(24).HeadingStyle:id=211,.parent=124"
      _StyleDefs(250) =   "Splits(1).Columns(24).FooterStyle:id=212,.parent=125"
      _StyleDefs(251) =   "Splits(1).Columns(24).EditorStyle:id=213,.parent=127"
      _StyleDefs(252) =   "Splits(1).Columns(25).Style:id=218,.parent=123,.locked=-1"
      _StyleDefs(253) =   "Splits(1).Columns(25).HeadingStyle:id=215,.parent=124"
      _StyleDefs(254) =   "Splits(1).Columns(25).FooterStyle:id=216,.parent=125"
      _StyleDefs(255) =   "Splits(1).Columns(25).EditorStyle:id=217,.parent=127"
      _StyleDefs(256) =   "Named:id=33:Normal"
      _StyleDefs(257) =   ":id=33,.parent=0"
      _StyleDefs(258) =   "Named:id=34:Heading"
      _StyleDefs(259) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(260) =   ":id=34,.wraptext=-1"
      _StyleDefs(261) =   "Named:id=35:Footing"
      _StyleDefs(262) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(263) =   "Named:id=36:Selected"
      _StyleDefs(264) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(265) =   "Named:id=37:Caption"
      _StyleDefs(266) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(267) =   "Named:id=38:HighlightRow"
      _StyleDefs(268) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(269) =   "Named:id=39:EvenRow"
      _StyleDefs(270) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(271) =   "Named:id=40:OddRow"
      _StyleDefs(272) =   ":id=40,.parent=33"
      _StyleDefs(273) =   "Named:id=41:RecordSelector"
      _StyleDefs(274) =   ":id=41,.parent=34"
      _StyleDefs(275) =   "Named:id=42:FilterBar"
      _StyleDefs(276) =   ":id=42,.parent=33"
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
      TabIndex        =   107
      Top             =   540
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.Label lbl_company 
      AutoSize        =   -1  'True
      Caption         =   "BRANCH OFFICE "
      Height          =   195
      Left            =   240
      TabIndex        =   89
      Top             =   180
      Width           =   1320
   End
End
Attribute VB_Name = "frm_mst_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim i_lang As Integer
Dim Click As Double, vEmployee_Code As String
Dim active As String
Dim access As String

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

str_sql = "select count(employee_code) as rec_count from m_employee where employee_code = '" _
& Replace$(Trim$(txt_employee_code), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Sub check_invalid()
MsgBox "Data found!", vbCritical, headerMSG
txt_employee_code = ""
If txt_employee_code.Enabled = True Then txt_employee_code.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If Not txt_employee_code = Adodc1.Recordset.Fields("employee_code").Value And _
check_validate_exist_new Then
    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

'validasi employee code
If Trim(txt_employee_code) = "" Then
    MsgBox "Employee Code is empty!", vbOKOnly + vbInformation, headerMSG
    txt_employee_code.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi employee name
If Trim(txt_employee_name) = "" Then
    MsgBox "Employee Name is empty!", vbOKOnly + vbInformation, headerMSG
    txt_employee_name.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi employee nick name
If Trim(txt_employee_nick_name) = "" Then
    MsgBox "Employee Nick Name is empty!", vbOKOnly + vbInformation, headerMSG
    txt_employee_nick_name.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi department tdbcombo
If check_validate_tdbcombo(TDBCombo_department) = False Then
    MsgBox "Department is not selected!", vbOKOnly + vbInformation, headerMSG
    TDBCombo_department.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi division tdbcombo
If check_validate_tdbcombo(TDBCombo_division) = False Then
    MsgBox "Division is not selected!", vbOKOnly + vbInformation, headerMSG
    TDBCombo_division.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi group shift tdbcombo
If check_validate_tdbcombo(TDBCombo_group_shift) = False Then
    MsgBox "Group Shift is not selected!", vbOKOnly + vbInformation, headerMSG
    TDBCombo_group_shift.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi last employment name
If Trim(txt_number_of_children) = "" Then
    MsgBox "Children is empty!", vbOKOnly + vbInformation, headerMSG
    txt_child_number.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi last employment name
If Trim(txt_npwp) = "" Then
    MsgBox "NPWP is empty!", vbOKOnly + vbInformation, headerMSG
    txt_npwp.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi last employment name
'If Trim(txt_last_employment_name) = "" Then
'    MsgBox "Last employment name is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_last_employment_name.SetFocus
'    check_validate_new = False
'    Exit Function
'End If

'validasi last employment title
'If Trim(txt_last_employment_title) = "" Then
'    MsgBox "Last employment title is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_last_employment_title.SetFocus
'    check_validate_new = False
'    Exit Function
'End If

'validasi place of birth
'If Trim(txt_place_of_birth) = "" Then
'    MsgBox "Place of birth is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_place_of_birth.SetFocus
'    check_validate_new = False
'    Exit Function
'End If

'validasi address
'If Trim(txt_address) = "" Then
'    MsgBox "Address is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_address.SetFocus
'    check_validate_new = False
'    Exit Function
'End If

'validasi phone number
'If Trim(txt_phone_number) = "" Then
'    MsgBox "Phone number is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_phone_number.SetFocus
'    check_validate_new = False
'    Exit Function
'End If

'validasi bank account
'If Trim(txt_bank_account) = "" Then
'    MsgBox "Bank account is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_bank_account.SetFocus
'    check_validate_new = False
'    Exit Function
'End If

'validasi last education code other
'If cbo_last_education_code.ListIndex = 0 Then
'If Trim(txt_last_education_code_other) = "" Then
'    MsgBox "Last education level is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_last_education_code_other.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
'End If

'validasi last education name
'If Trim(txt_last_education_name) = "" Then
'    MsgBox "Last education name is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_last_education_name.SetFocus
'    check_validate_new = False
'    Exit Function
'End If

'validasi Title
If check_validate_tdbcombo(TDBCombo_title) = False Then
    MsgBox "Title is not selected!", vbOKOnly + vbInformation, headerMSG
    TDBCombo_title.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi descripton
'If Trim(txt_description) = "" Then
'    MsgBox "Description is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_description.SetFocus
'    check_validate_new = False
'    Exit Function
'End If

'validasi not active-reason
'If opt_not_active Then
'If Trim(txt_end_working_reason) = "" Then
'    MsgBox "Reason is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_end_working_reason.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
'End If
End Function

Private Sub btnChange_Click()
TDBCombo_company1.RowSource = Adodc_company
    Click = Click + 1
    If Click Mod 2 = 1 Then
        btnChange.Caption = "Cancel"
        TDBCombo_company1.Visible = True
        txt_company_name1.Visible = True
        
        TDBCombo_department.Text = ""
        txt_department_name.Text = ""
        TDBCombo_division.Text = ""
        txt_division_name.Text = ""
    Else
        btnChange.Caption = "Change Branch Office"
        TDBCombo_company1.Visible = False
        txt_company_name1.Visible = False
    End If
End Sub

Private Sub cbo_last_education_code_Click()
If cbo_last_education_code.ListIndex = 0 Then
    txt_last_education_code_other.Visible = True
Else
    txt_last_education_code_other = "": txt_last_education_code_other.Visible = False
End If
End Sub

Private Sub cmd_refresh_Click()
Call load_data_employee
End Sub

Private Sub cmd_import_Click()
'If check_validate_tdbcombo(TDBCombo_company) = False Then
'    MsgBox "Company is not selected!", vbInformation, headerMSG
'    Exit Sub
'End If

frm_trans_import_employee.Show 1
End Sub

Private Sub cmd_print_Click()
'Dim str_file As String
'Dim a As New frm_rpt
'Dim strsql As String, active As String
'
'str_file = "\report\rpt_master_employee.rpt"
'
'active = IIf(optActive.Value, "<> 0", "= 0")
'access = IIf(LOGIN_LEVEL = 100, "", "AND (a.managerial_access = 0 OR a.managerial_access IS NULL)")
'
'strsql = "SELECT a.company_code,b.company_name,a.department_code,c.department_name," _
'            & "a.employee_code,a.employee_name,a.title_code,e.title_name,a.marital_status," _
'            & "a.sex,a.start_working,a.bank_account,a.branches,a.flag_active,a.division_code," _
'            & "d.division_name,a.number_of_children,a.religion,a.npwp,a.phone_number," _
'            & "a.place_of_birth,a.date_of_birth,a.address " _
'            & "FROM m_employee a LEFT JOIN m_company b ON a.company_code = b.company_code " _
'            & "LEFT JOIN m_department c ON a.department_code = c.department_code AND c.company_code = b.company_code " _
'            & "LEFT JOIN m_division d ON a.division_code = d.division_code AND d.department_code = c.department_code AND d.company_code = b.company_code " _
'            & "LEFT JOIN m_title e ON a.title_code = e.title_code " _
'            & "where a.company_code = '" _
'            & TDBCombo_company.Columns("company_code").Value & "' " _
'            & "AND a.flag_active " & active & " " & access & " order by a.employee_code"
'
'Call a.Show
'
'a.Caption = "REPORT MASTER EMPLOYEE"
'Call a.rpt_view(strsql, str_file, Now)

frm_lookup_sort_employee.Show

End Sub

Private Sub cmd_update_Click()
TDBGrid1.Update
CnG.Execute "update m_employee set flag_active=1 where flag_active=-1"
End Sub

Private Sub CmdCancel_Click()
int_mode = 0
Call load_mode

btnChange.Visible = False
TDBCombo_company1.Visible = False
txt_company_name1.Visible = False

btnChange.Caption = "Change Branch Office"
Frame3.Visible = True
Click = 0
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "No Data selected!", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Are you sure want to delete data '" _
    & TDBGrid1.Columns("employee_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from m_employee where employee_code = '" _
    & TDBGrid1.Columns("employee_code").Value & "' AND " _
    & "company_code = '" & TDBGrid1.Columns("company_code").Value & "' AND " _
    & "department_code = '" & TDBGrid1.Columns("department_code").Value & "' AND " _
    & "division_code = '" & TDBGrid1.Columns("division_code").Value & "'"
CnG.CommitTrans

Call load_data_employee
int_mode = 0
Call load_mode
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

Private Sub set_data_division(ByRef str_code As String)
On Error Resume Next

Adodc_division.Recordset.MoveFirst
Adodc_division.Recordset.Find ("division_code='" & str_code & "'")   ', 0, adSearchForward, 1)
If Not (Adodc_division.Recordset.EOF = True Or Adodc_division.Recordset.BOF = True) Then
    TDBCombo_division.Bookmark = Adodc_division.Recordset.AbsolutePosition
    Call TDBCombo_division_ItemChange
Else
    TDBCombo_division.Text = ""
End If
End Sub

Private Sub set_data_title(ByVal str_code As String)
On Error Resume Next

Adodc_title.Recordset.MoveFirst
Adodc_title.Recordset.Find ("title_code='" & str_code & "'")   ', 0, adSearchForward, 1)
If Not (Adodc_title.Recordset.EOF = True Or Adodc_title.Recordset.BOF = True) Then
    TDBCombo_title.Bookmark = Adodc_title.Recordset.AbsolutePosition
    Call TDBCombo_title_ItemChange
Else
    TDBCombo_title.Text = ""
End If
End Sub

Private Sub set_data_group_shift(ByVal str_code As String)
On Error Resume Next

Adodc_group_shift.Recordset.MoveFirst
Adodc_group_shift.Recordset.Find ("group_code='" & str_code & "'")   ', 0, adSearchForward, 1)
If Not (Adodc_group_shift.Recordset.EOF = True Or Adodc_group_shift.Recordset.BOF = True) Then
    TDBCombo_group_shift.Bookmark = Adodc_group_shift.Recordset.AbsolutePosition
    Call TDBCombo_group_shift_ItemChange
Else
    TDBCombo_group_shift.Text = ""
End If
End Sub

Private Function get_data_obj(ByRef Ctr As Control, ByVal str As Variant) As Variant
If TypeOf Ctr Is ComboBox Then
    If Ctr.name = "cbo_sex" Or Ctr.name = "cbo_marital_status" Or Ctr.name = "cbo_religion" Then
        get_data_obj = IIf(IsNull(str) = True, 1, str)
    
    ElseIf Ctr.name = "cbo_last_education_code" Then
        Ctr.Text = IIf(IsNull(str), "SMA", str)
        'get_data_obj = IIf(IsNull(str) = True, 1, str)
    End If
    
'    If Ctr.name = "cbo_sex" Then
'        get_data_obj = IIf(IsNull(str) = True, 1, str)
'    Else
'        get_data_obj = IIf(IsNull(str) = True, 0, str)
'    End If
    
ElseIf TypeOf Ctr Is DTPicker Then
    get_data_obj = IIf(IsNull(str) = True, Now, str)
ElseIf TypeOf Ctr Is TextBox Then
    get_data_obj = IIf(IsNull(str) = True, "", str)
End If
End Function

Public Sub set_edit_data()
Dim kddivisi As String
Dim v_end_mc As String
With Adodc1.Recordset
    
    Call load_data_division(.Fields("department_code"))
    
    txt_employee_code = .Fields("employee_code").Value
    txt_branch = IIf(IsNull(.Fields("branches").Value), "", IsNull(.Fields("branches").Value))
    '-----------------------------------------------------------
    txt_employee_name = .Fields("employee_name").Value
    txt_employee_nick_name = "" & .Fields("employee_nick_name").Value
    
    Call set_data_department(.Fields("department_code").Value)
    Call set_data_group_shift(IIf(IsNull(.Fields("group_shift").Value), "", .Fields("group_shift").Value))
    kddivisi = .Fields("division_code").Value
    Call set_data_division(kddivisi)
'    Call load_data_department
    
    DTPicker_birth.Value = get_data_obj(DTPicker_birth, .Fields("date_of_birth").Value)
    Call set_age_data
    txt_place_of_birth = get_data_obj(txt_place_of_birth, .Fields("place_of_birth").Value)
    cbo_sex.ListIndex = get_data_obj(cbo_sex, .Fields("sex").Value)
    cbo_religion.ListIndex = get_data_obj(cbo_religion, .Fields("religion").Value)
    cbo_marital_status.ListIndex = get_data_obj(cbo_marital_status, .Fields("marital_status").Value)
    
    txt_number_of_children = Val("" & .Fields("number_of_children").Value)
    txt_address = get_data_obj(txt_address, .Fields("address").Value)
    txt_email = "" & .Fields("email").Value
    txt_npwp = "" & .Fields("npwp").Value
    
    txt_phone_number = get_data_obj(txt_phone_number, .Fields("phone_number").Value)
    txt_bank_account = get_data_obj(txt_bank_account, .Fields("bank_account").Value)
    cbo_last_education_code.ListIndex = get_data_obj(cbo_last_education_code, .Fields("last_education_code").Value)
    
    If cbo_last_education_code.ListIndex = 0 Then
        txt_last_education_code_other = get_data_obj(txt_last_education_code_other, .Fields("last_education_code_other").Value)
    End If
    txt_last_education_name = get_data_obj(txt_last_education_name, .Fields("last_education_name").Value)
    DTPicker_last_education_pass.Value = get_data_obj(DTPicker_last_education_pass, .Fields("last_education_pass").Value)
    txt_last_employment_name = get_data_obj(txt_last_employment_name, .Fields("last_employment_name").Value)
    DTPicker_last_employment.Value = get_data_obj(DTPicker_last_employment, .Fields("last_employment_date").Value)
    txt_last_employment_title = get_data_obj(txt_last_employment_title, .Fields("last_employment_title").Value)
    DTPicker_start_working.Value = get_data_obj(DTPicker_start_working, .Fields("start_working").Value)
    DTPicker_appointment.Value = get_data_obj(DTPicker_appointment, .Fields("date_of_appointment").Value)
    
    Call set_data_title(.Fields("title_code").Value)
    
    cbo_shiftable.ListIndex = Abs(get_data_obj(cbo_shiftable, .Fields("flag_shiftable").Value))
    txt_description = get_data_obj(txt_description, .Fields("description").Value)
    
'    opt_active.Value = IIf(.Fields("flag_active").Value = 1, 1, 0)
'    opt_not_active.Value = IIf(.Fields("flag_active").Value = 0, 1, 0)
'    If opt_not_active Then
'        DTPicker_end_working.Value = get_data_obj(DTPicker_end_working, .Fields("end_working").Value)
'        txt_end_working_reason = get_data_obj(txt_end_working_reason, .Fields("reason").Value)
'    End If
    opt_active.Value = IIf(.Fields("flag_active").Value = 1, 1, 0)
    opt_not_active.Value = IIf(.Fields("flag_active").Value = 0, 1, 0)
    opt_mc.Value = IIf(.Fields("flag_active").Value = 2, 1, 0)
    
'    v_end_mc = IIf(IsNull(.Fields("end_mc")), "0000:00:00", DateValue(.Fields("end_mc")))
    If Not .Fields("end_mc") = "" Then
        lbl_end_mc.Visible = True
        dt_end_mc.Visible = True
        dt_end_mc.Value = get_data_obj(dt_end_mc, .Fields("end_mc").Value)
    Else
        Label12.Visible = False
        dt_end_mc.Visible = False
    End If
    
    If opt_not_active.Value Then
        DTPicker_end_working.Value = get_data_obj(DTPicker_end_working, .Fields("end_working").Value)
        txt_end_working_reason = get_data_obj(txt_end_working_reason, .Fields("reason").Value)
    End If
    
    If opt_mc.Value Then
        DTPicker_end_working.Value = get_data_obj(DTPicker_end_working, .Fields("start_mc").Value)
        txt_end_working_reason = get_data_obj(txt_end_working_reason, .Fields("reason").Value)
    End If
    
    Check1.Value = IIf(IsNull(.Fields("managerial_level").Value) = True, 0, 1)
    
    Call set_age_data
    Call set_working_age_data
    
    txt_fathers_name = get_data_obj(txt_fathers_name, .Fields("fathers_name").Value)
    txt_mothers_name = get_data_obj(txt_mothers_name, .Fields("mothers_name").Value)
    txt_child_number = get_data_obj(txt_child_number, .Fields("child_number").Value)
    txt_child_number_from = get_data_obj(txt_child_number_from, .Fields("child_number_from").Value)
    Check1.Value = IIf(IsNull(.Fields("managerial_level").Value), 0, .Fields("managerial_level").Value)
    Check2.Value = IIf(IsNull(.Fields("managerial_access").Value), 0, .Fields("managerial_access").Value)

    txt_npwp_address = get_data_obj(txt_npwp_address, .Fields("npwp_address").Value)
End With
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close
    rsBound.Open "select * from m_employee where employee_code = '" _
    & Adodc1.Recordset.Fields("employee_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic
rsBound.Close

int_mode = 2
Click = 0
Frame3.Visible = False

Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_employee where employee_code = 'άφ'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 1
Frame3.Visible = False
Call load_mode
End Sub

Private Sub CmdPrint_Click()
TDBGrid1.PrintInfo.PageSetup
If Not TDBGrid1.PrintInfo.PageSetupCancelled = True Then
    TDBGrid1.PrintInfo.PrintPreview dbgAllRows
End If
End Sub

Private Sub insert_new_data()
On Error GoTo err

CnG.BeginTrans

With rsBound
    .AddNew
    
    
    .Fields("employee_code").Value = Trim(txt_employee_code)
    '-----------------------------------------------------------------------------
    .Fields("employee_name").Value = Replace(Trim(txt_employee_name), "'", "''")
    .Fields("employee_nick_name").Value = Trim(txt_employee_nick_name)
    
    .Fields("company_code").Value = TDBCombo_company.Columns("company_code").Value
    .Fields("company_name").Value = Trim(txt_company_name)
    .Fields("company_name").Value = Trim(txt_company_name)
    .Fields("department_code").Value = TDBCombo_department.Columns("department_code").Value
    .Fields("department_name").Value = Trim(txt_department_name)
    .Fields("division_code").Value = TDBCombo_division.Columns("division_code").Value
    .Fields("division_name").Value = Trim(txt_division_name)
    
    .Fields("date_of_birth").Value = Format(DTPicker_birth.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("place_of_birth").Value = Trim(txt_place_of_birth)
    .Fields("sex").Value = cbo_sex.ListIndex
    .Fields("religion").Value = cbo_religion.ListIndex
    .Fields("marital_status").Value = cbo_marital_status.ListIndex
    .Fields("number_of_children").Value = Val(DropAllComma(txt_number_of_children))
    .Fields("address").Value = Trim(txt_address)
    .Fields("email").Value = Trim(txt_email)
    .Fields("npwp").Value = Trim(txt_npwp)
    
    .Fields("phone_number").Value = Trim(txt_phone_number)
    .Fields("bank_account").Value = Trim(txt_bank_account)
    .Fields("last_education_code").Value = cbo_last_education_code.ListIndex
    If cbo_last_education_code.ListIndex = 0 Then
        .Fields("last_education_code_other").Value = Trim(txt_last_education_code_other)
    End If
    .Fields("last_education_name").Value = Trim(txt_last_education_name)
    .Fields("last_education_pass").Value = Format(DTPicker_last_education_pass.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("last_employment_name").Value = Trim(txt_last_employment_name)
    .Fields("last_employment_date").Value = Format(DTPicker_last_employment.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("last_employment_title").Value = Trim(txt_last_employment_title)
    .Fields("start_working").Value = Format(DTPicker_start_working.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("date_of_appointment").Value = Format(DTPicker_appointment.Value, "yyyy-MM-dd HH:mm:ss")
    
    .Fields("title_code").Value = Trim(txt_title_name)
    .Fields("title_name").Value = TDBCombo_title.Columns("title_name").Value
    .Fields("flag_shiftable").Value = cbo_shiftable.ListIndex
    .Fields("description").Value = Trim(txt_description)
    
    .Fields("flag_active").Value = IIf(opt_active, 1, 0)
    If opt_not_active Then
        .Fields("end_working").Value = Format(DTPicker_end_working.Value, "yyyy-MM-dd HH:mm:ss")
        .Fields("reason").Value = Trim(txt_end_working_reason)
    End If
    
    .Fields("fathers_name").Value = Trim(txt_fathers_name)
    .Fields("mothers_name").Value = Trim(txt_mothers_name)
    .Fields("child_number").Value = Val(txt_child_number)
    .Fields("child_number_from").Value = Val(txt_child_number_from)
    .Fields("managerial_level").Value = Check1.Value
    .Fields("managerial_access").Value = Check2.Value
    
    .Fields("branches").Value = Trim(txt_branch)
    .Fields("group_shift").Value = Trim(TDBCombo_group_shift.Text)
    .Fields("npwp_address").Value = Trim(txt_npwp_address.Text)
    
    .Update
End With

CnG.CommitTrans
Exit Sub

err:
MsgBox "There Is Any Problem With Application!" & Chr(13) & _
    "Please Contact Us (PT. Solusi Sentral Data - (031) 5616465)", vbInformation, headerMSG
Exit Sub
End Sub

Private Sub edit_old_data()
Dim strsql As String

On Error GoTo err_capture

'CnG.BeginTrans
'With rsBound
'
'    .Fields("employee_code").Value = Trim(txt_employee_code)
'    '-----------------------------------------------------------------------------
'    .Fields("employee_name").Value = Trim(txt_employee_name)
'    .Fields("employee_nick_name").Value = Trim(txt_employee_nick_name)
'
'    .Fields("company_code").Value = TDBCombo_company.Columns("company_code").Value
'    .Fields("company_name").Value = Trim(txt_company_name)
'    .Fields("department_code").Value = TDBCombo_department.Columns("department_code").Value
'    .Fields("department_name").Value = Trim(txt_department_name)
'    .Fields("division_code").Value = TDBCombo_division.Columns("division_code").Value
'    .Fields("division_name").Value = Trim(txt_division_name)
'
'    .Fields("date_of_birth").Value = Format(DTPicker_birth.Value, "yyyy-MM-dd HH:mm:ss")
'    .Fields("place_of_birth").Value = Trim(txt_place_of_birth)
'    .Fields("sex").Value = cbo_sex.ListIndex
'    .Fields("religion").Value = cbo_religion.ListIndex
'    .Fields("marital_status").Value = cbo_marital_status.ListIndex
'    .Fields("number_of_children").Value = Val(DropAllComma(txt_number_of_children))
'    .Fields("address").Value = Trim(txt_address)
'    .Fields("email").Value = Trim(txt_email)
'    .Fields("npwp").Value = Trim(txt_npwp)
'
'    .Fields("phone_number").Value = Trim(txt_phone_number)
'    .Fields("bank_account").Value = Trim(txt_bank_account)
'    .Fields("last_education_code").Value = cbo_last_education_code.ListIndex
'    If cbo_last_education_code.ListIndex = 0 Then
'        .Fields("last_education_code_other").Value = Trim(txt_last_education_code_other)
'    End If
'    .Fields("last_education_name").Value = Trim(txt_last_education_name)
'    .Fields("last_education_pass").Value = Format(DTPicker_last_education_pass.Value, "yyyy-MM-dd HH:mm:ss")
'    .Fields("last_employment_name").Value = Trim(txt_last_employment_name)
'    .Fields("last_employment_date").Value = Format(DTPicker_last_employment.Value, "yyyy-MM-dd HH:mm:ss")
'    .Fields("last_employment_title").Value = Trim(txt_last_employment_title)
'    .Fields("start_working").Value = Format(DTPicker_start_working.Value, "yyyy-MM-dd HH:mm:ss")
'    .Fields("date_of_appointment").Value = Format(DTPicker_appointment.Value, "yyyy-MM-dd HH:mm:ss")
'
'    .Fields("title_code").Value = TDBCombo_title.Columns("title_code").Value
'    .Fields("title_name").Value = Trim(txt_title_name)
'    .Fields("flag_shiftable").Value = cbo_shiftable.ListIndex
'    .Fields("description").Value = Trim(txt_description)
'
'    .Fields("flag_active").Value = IIf(opt_active, 1, 0)
'    If opt_not_active Then
'        .Fields("end_working").Value = Format(DTPicker_end_working.Value, "yyyy-MM-dd HH:mm:ss")
'        .Fields("reason").Value = Trim(txt_end_working_reason)
'    End If
'
'    .Fields("fathers_name").Value = Trim(txt_fathers_name)
'    .Fields("mothers_name").Value = Trim(txt_mothers_name)
'    .Fields("child_number").Value = Val(txt_child_number)
'    .Fields("child_number_from").Value = Val(txt_child_number_from)
'    .Fields("managerial_level").Value = Check1.Value
'
'    .Update
'End With
'CnG.CommitTrans

'If TDBCombo_department = "" Then
'    MsgBox "Data Department tidak boleh kosong !"
'    Call set_data_mode
'    TDBCombo_department.SetFocus
'    Exit Sub
'End If
Dim rscari As New ADODB.Recordset
Dim tglAkhirAbsen As Date

    CnG.BeginTrans
    
    If Adodc1.Recordset.Fields("flag_active") = 2 Then
        strsql = "SELECT employee_code, att_date from h_attendance where employee_code = '" & txt_employee_code.Text & "' " & _
            "order by att_date DESC LIMIT 1"
        rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If Not rscari.EOF Then
            tglAkhirAbsen = DateValue(rscari!att_date)
        
        
        While tglAkhirAbsen < DateValue(dt_end_mc.Value)
            tglAkhirAbsen = tglAkhirAbsen + 1
'            strSQL = "INSERT INTO h_attendance ........"
            strsql = "INSERT INTO h_attendance (employee_code, att_date," & _
                        "shift_number, shift_code, start_time," & _
                        "end_time," & _
                        "flag_present,description, entry_date) " & _
                    "VALUES " & _
                        "('" & rscari!employee_code & "','" & Format(tglAkhirAbsen, "yyyy-MM-dd 00:00:00") & "'," & _
                        "1, 'MC','" & Format(tglAkhirAbsen, "yyyy-MM-dd 08:00:00") & "'," & _
                        "'" & Format(tglAkhirAbsen, "yyyy-MM-dd 17:00:00") & "'," & _
                            "0,'MC', Now())"
            CnG.Execute strsql
            
        Wend
    End If
    End If
    
    If TDBCombo_company1.Visible = True Then
        strsql = "UPDATE m_employee set employee_code = '" & Trim(txt_employee_code.Text) & "',employee_name = '" & Replace(Trim(txt_employee_name), "'", "''") & "'," _
            & "employee_nick_name = '" & Replace(Trim(txt_employee_nick_name), "'", "''") & "',company_code = '" & TDBCombo_company1.Columns("company_code").Value & "', " _
            & "company_name = '" & Replace(Trim(txt_company_name1), "'", "''") & "',department_code = '" & TDBCombo_department.Columns("department_code").Value & "',"
    Else
        strsql = "UPDATE m_employee set employee_code = '" & Trim(txt_employee_code.Text) & "',employee_name = '" & Replace(Trim(txt_employee_name), "'", "''") & "'," _
            & "employee_nick_name = '" & Replace(Trim(txt_employee_nick_name), "'", "''") & "',company_code = '" & TDBCombo_company.Columns("company_code").Value & "', " _
            & "company_name = '" & Replace(Trim(txt_company_name), "'", "''") & "',department_code = '" & IIf(IsNull(TDBCombo_department.Columns("department_code").Value), "N/A", TDBCombo_department.Columns("department_code").Value) & "',"
    End If
    
    strsql = strsql _
            & "department_name = '" & Replace(Trim(txt_department_name), "'", "''") & "',division_code = '" & TDBCombo_division.Columns("division_code").Value & "'," _
            & "division_name = '" & Replace(Trim(txt_division_name), "'", "''") & "',date_of_birth = '" & Format(DTPicker_birth.Value, "yyyy-MM-dd HH:mm:ss") & "'," _
            & "place_of_birth = '" & Replace(Trim(txt_place_of_birth), "'", "''") & "',sex = '" & cbo_sex.ListIndex & "',religion = '" & cbo_religion.ListIndex & "'," _
            & "marital_status = '" & cbo_marital_status.ListIndex & "',number_of_children = '" & Val(DropAllComma(txt_number_of_children)) & "'," _
            & "address = '" & Trim(txt_address) & "',email = '" & Trim(txt_email) & "',npwp = '" & Trim(txt_npwp) & "',phone_number = '" & Trim(txt_phone_number) & "'," _
            & "bank_account = '" & Trim(txt_bank_account) & "',last_education_code = '" & cbo_last_education_code.ListIndex & "'," _
            & "last_education_code_other = '" & IIf(cbo_last_education_code.ListIndex = 0, Trim(txt_last_education_code_other), Null) & "'," _
            & "last_education_name = '" & Trim(txt_last_education_name) & "',last_education_pass = '" & Format(DTPicker_last_education_pass.Value, "yyyy-MM-dd HH:mm:ss") & "'," _
            & "last_employment_name = '" & Trim(txt_last_employment_name) & "',last_employment_date = '" & Format(DTPicker_last_employment.Value, "yyyy-MM-dd HH:mm:ss") & "'," _
            & "last_employment_title = '" & Trim(txt_last_employment_title) & "',start_working = '" & Format(DTPicker_start_working.Value, "yyyy-MM-dd HH:mm:ss") & "'," _
            & "date_of_appointment = '" & Format(DTPicker_appointment.Value, "yyyy-MM-dd HH:mm:ss") & "',title_code = '" & Trim(txt_title_name) & "'," _
            & "title_name = '" & TDBCombo_title.Columns("title_name").Value & "',flag_shiftable = '" & cbo_shiftable.ListIndex & "',description = '" & Trim(txt_description) & "'," _
            & "flag_active = '" & IIf(opt_active.Value, 1, IIf(opt_not_active.Value, 0, 2)) & "'," _
            & "end_working = '" & IIf(opt_not_active.Value, Format(DTPicker_end_working.Value, "yyyy-MM-dd HH:mm:ss"), "00:00:00") & "', " _
            & "reason = '" & IIf(opt_not_active.Value Or opt_mc.Value, Trim(txt_end_working_reason), "") & "', " _
            & "fathers_name = '" & Trim(txt_fathers_name) & "',mothers_name = '" & Trim(txt_mothers_name) & "',child_number = '" & Val(txt_child_number) & "', " _
            & "child_number_from = '" & Val(txt_child_number_from) & "',managerial_level = '" & Check1.Value & "', " _
            & "start_mc = '" & IIf(opt_mc.Value, Format(DTPicker_end_working, "yyyy-MM-dd"), IIf(opt_active.Value And dt_end_mc.Visible = True, Format(DTPicker_end_working, "yyyy-MM-dd"), "00:00:00")) & "', " _
            & "end_mc = '" & IIf(opt_active.Value And dt_end_mc.Visible = True, Format(dt_end_mc, "yyyy-MM-dd"), "0000:00:00") & "', " _
            & "branches = '" & Trim(txt_branch) & "',managerial_access = '" & Check2.Value & "',group_shift = '" & Trim(TDBCombo_group_shift.Text) & "', " _
            & "npwp_address = '" & Trim(txt_npwp_address.Text) & "' " _
            & "WHERE employee_code = '" & Adodc1.Recordset.Fields("employee_code").Value & "'"
        CnG.Execute strsql
        
    CnG.CommitTrans
Exit Sub
err_capture:
rsBound.CancelBatch adAffectCurrent: rsBound.Close: CnG.RollbackTrans
End Sub

Private Sub CmdSave_Click()
On Error Resume Next

Dim strsql As String
Dim int_proses As Integer, vli_Loop As Integer

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
    
    '++++++++++++ Cek Data Employee Code +++++++++++++++++
    If vEmployee_Code <> txt_employee_code Then
        int_proses = MsgBox("Terdapat Perubahan No. Badge." & Chr(13) & _
            "Apakah Anda akan Melanjutkan Proses?", vbYesNo, headerMSG)
        If Not int_proses = vbYes Then Exit Sub
        
        frmLoading.Visible = True
        
        CnG.BeginTrans
        
        Dim vTabel As String
        For vli_Loop = 0 To 29
            vTabel = IIf(vli_Loop = 0, "h_attendance", IIf(vli_Loop = 1, "h_d_salary", IIf(vli_Loop = 2, "h_log_attendance", _
                    IIf(vli_Loop = 3, "h_log_attendance_recover", IIf(vli_Loop = 4, "h_salary", IIf(vli_Loop = 5, "h_salary_eks", _
                    IIf(vli_Loop = 6, "h_salary_new", IIf(vli_Loop = 7, "h_send_mail", IIf(vli_Loop = 8, "h_spt21", _
                    IIf(vli_Loop = 9, "m_enroll_link", IIf(vli_Loop = 10, "m_salary", IIf(vli_Loop = 11, "t_absent", _
                    IIf(vli_Loop = 12, "t_check", IIf(vli_Loop = 13, "t_duty", IIf(vli_Loop = 14, "t_employee_expense", _
                    IIf(vli_Loop = 15, "t_employee_income", IIf(vli_Loop = 16, "t_employee_performance", _
                    IIf(vli_Loop = 17, "t_leave", IIf(vli_Loop = 18, "t_leave_periode", IIf(vli_Loop = 19, "t_loan", _
                    IIf(vli_Loop = 20, "t_log_attendance", IIf(vli_Loop = 21, "t_pensiun", IIf(vli_Loop = 22, "t_salary_item", _
                    IIf(vli_Loop = 23, "t_subsidy", IIf(vli_Loop = 24, "t_thr", IIf(vli_Loop = 25, "td_claim", _
                    IIf(vli_Loop = 26, "td_loan", IIf(vli_Loop = 27, "td_shift", IIf(vli_Loop = 28, "tm_claim", "tm_loan")))))))))))))))))))))))))))))
            
            strsql = "UPDATE " & vTabel & " SET employee_code = '" & txt_employee_code.Text & "' " _
                    & "WHERE employee_code = '" & vEmployee_Code & "'"
            CnG.Execute strsql
        Next vli_Loop
        
        CnG.CommitTrans
        frmLoading.Visible = False
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++

    If opt_mc = True And DTPicker_end_working.Value < Now() Then
        int_proses = MsgBox("Data Kehadiran Setelah Tanggal Start MC akan Dihapus" & Chr(13) & _
            "Apakah Anda akan Melanjutkan Proses?", vbYesNo, headerMSG)
        If int_proses = vbYes Then
            strsql = "DELETE FROM h_attendance " & _
                    "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                    "AND DATE(att_date) >= '" & Format(DTPicker_end_working.Value, "yyyy-MM-dd") & "'"
            CnG.Execute strsql
            MsgBox "Attendance Data Succesfully Deleted"
        Else
            MsgBox "Delete Data Canceled"
            Call load_data_employee
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
    End If
  
    Call edit_old_data
End If

Call load_data_employee
int_mode = 0
Call load_mode

TDBCombo_company1.Visible = False
txt_company_name1.Visible = False
btnChange.Caption = "Change Branch Office"
btnChange.Visible = False
Frame3.Visible = True
Click = 0
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal F As Boolean, ByVal g As Boolean)
CmdNew.Enabled = a And blnUser_Add
CmdSave.Enabled = b
cmdEdit.Enabled = c And blnUser_Edit
cmdDelete.Enabled = d And blnUser_Delete
CmdCancel.Enabled = e

CmdPrint.Enabled = F
'cmd_refresh.Enabled = g
cmd_import.Enabled = blnUser_Add
End Sub

Private Sub clear_view_data()
Dim Ctr As Control
For Each Ctr In Me
    If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
        If Not LCase(Ctr.name) = "txt_company_name" Then Ctr.Text = ""
    ElseIf TypeOf Ctr Is TDBCombo Then
        If Not LCase(Ctr.name) = "tdbcombo_company" Then Ctr.Text = ""
    ElseIf TypeOf Ctr Is DTPicker Then
        Ctr.Value = Now
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
cbo_sex.ListIndex = 1
cbo_religion.ListIndex = 1
cbo_marital_status.ListIndex = 0
cbo_last_education_code.ListIndex = 3
cbo_shiftable.ListIndex = 0

DTPicker_appointment.Value = Now
DTPicker_birth.Value = Now
DTPicker_end_working.Value = Now
DTPicker_last_education_pass.Value = Now
DTPicker_last_employment.Value = Now
DTPicker_start_working.Value = Now

Label12.Visible = False
dt_end_mc.Visible = False

Call set_age_data
Call set_working_age_data
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
    txt_employee_code.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If txt_employee_code.Enabled = True Then
        txt_employee_code.SetFocus
    End If
    vEmployee_Code = ""
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    TDBGrid1.Enabled = True
    vEmployee_Code = ""

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
'    txt_employee_code.Enabled = False
    fra_entry.Visible = True
    btnChange.Visible = True
    TDBGrid1.Enabled = False
    vEmployee_Code = txt_employee_code.Text
    
    '++++++++++ Chek Managerial Access ++++++++++++++
    If LOGIN_LEVEL = 100 Then
        Check2.Visible = True
    Else
        Check2.Visible = False
    End If
    '++++++++++++++++++++++++++++++++++++++++++++++++
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

'Private Sub cmd_print_Click()
'MsgBox Chr$(40)
'End Sub

Private Sub set_age_data()
txt_age = Trim(str(Year(Now) - Year(DTPicker_birth.Value)))
End Sub

Private Sub set_working_age_data()
txt_working_time = Trim(str(Year(Now) - Year(DTPicker_start_working.Value)))
End Sub

Private Sub DTPicker_birth_Change()
Call set_age_data
End Sub

Private Sub DTPicker_start_working_Change()
Call set_working_age_data
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = strConn
Adodc_company.ConnectionString = strConn
Adodc_department.ConnectionString = strConn
Adodc_division.ConnectionString = strConn
Adodc_title.ConnectionString = strConn
Adodc_group_shift.ConnectionString = strConn

btnChange.Visible = False
TDBCombo_company1.Visible = False
txt_company_name1.Visible = False

lbl_end_mc.Visible = False
dt_end_mc.Visible = False

optActive.Value = True

Call load_data_company
Call load_data_title

Call load_data_user_access(Me)
int_mode = 0
Call load_mode
timer1.Enabled = True
Frame3.Visible = True

'Call set_language

cmdEdit.Enabled = False
cmdDelete.Enabled = False
CmdNew.Enabled = False
Frame3.Enabled = False

End Sub

Private Sub set_language()
Dim i_lbl_company(2), i_lbl_employee_code(2), i_lbl_employee_name(2), i_lbl_employee_nick_name(2) As String
Dim i_lbl_department(2), i_lbl_division(2), i_fra_last_employment(2), i_lbl_last_employment_name(2), i_lbl_last_employment_date(2), i_lbl_last_employment_title(2) As String
Dim i_lbl_date_of_birth(2), i_lbl_age(2), i_lbl_place_of_birth(2), i_lbl_sex(2), i_lbl_religion(2), i_lbl_marital_status(2), i_lbl_address(2), i_lbl_phone_number(2), i_lbl_bank_account(2) As String
Dim i_fra_last_education(2), i_lbl_last_education_level(2), i_lbl_last_education_name(2), i_lbl_last_education_pass(2) As String
Dim i_lbl_start_working(2), i_lbl_appointment(2), i_lbl_working_age(2), i_lbl_title(2), i_lbl_description(2), i_lbl_shiftable(2), i_lbl_reason_end_working(2), i_lbl_reason_description(2) As String

i_lbl_company(0) = "BRANCH OFFICE": i_lbl_company(1) = "KANTOR CABANG"
i_lbl_employee_code(0) = "EMP. CODE": i_lbl_employee_code(1) = "N I P"
i_lbl_employee_name(0) = "EMP. NAME": i_lbl_employee_name(1) = "NAMA"
i_lbl_employee_nick_name(0) = "NICK NAME": i_lbl_employee_nick_name(1) = "PANGGILAN"
i_lbl_department(0) = "DEPARTMENT": i_lbl_department(1) = "DEPARTEMEN"
i_lbl_division(0) = "DIVISION": i_lbl_division(1) = "DIVISI"
i_fra_last_employment(0) = "LAST EMPLOYMENT": i_fra_last_employment(1) = "PEKERJAAN TERAKHIR"
i_lbl_last_employment_name(0) = "NAME": i_lbl_last_employment_name(1) = "NAMA"
i_lbl_last_employment_date(0) = "DATE": i_lbl_last_employment_date(1) = "TGL"
i_lbl_last_employment_title(0) = "TITLE": i_lbl_last_employment_title(1) = "JAB."
i_lbl_date_of_birth(0) = "DATE OF BIRTH": i_lbl_date_of_birth(1) = "TGL. LAHIR"
i_lbl_age(0) = "AGE (Y)*": i_lbl_age(1) = "UMUR"
i_lbl_place_of_birth(0) = "PLACE OF BIRTH": i_lbl_place_of_birth(1) = "TMP. LAHIR"
i_lbl_sex(0) = "SEX": i_lbl_sex(1) = "J. KELAMIN"
i_lbl_religion(0) = "RELIGION": i_lbl_religion(1) = "AGAMA"
i_lbl_marital_status(0) = "MARITAL STATUS": i_lbl_marital_status(1) = "S. PERKAWINAN"
i_lbl_address(0) = "ADDRESS": i_lbl_address(1) = "ALAMAT"
i_lbl_phone_number(0) = "PHONE NUMBER": i_lbl_phone_number(1) = "NO. TELP."
i_lbl_bank_account(0) = "BANK ACCOUNT": i_lbl_bank_account(1) = "REK. BANK"
i_fra_last_education(0) = "LAST EDUCATION": i_fra_last_education(1) = "PEND. TERAKHIR"
i_lbl_last_education_level(0) = "LEVEL": i_lbl_last_education_level(1) = "LEVEL"
i_lbl_last_education_name(0) = "NAME": i_lbl_last_education_name(1) = "NAMA"
i_lbl_last_education_pass(0) = "PASS": i_lbl_last_education_pass(1) = "TAHUN"
i_lbl_start_working(0) = "START WORKING": i_lbl_start_working(1) = "MULAI BEKERJA"
i_lbl_appointment(0) = "APPOINTMENT": i_lbl_appointment(1) = "PENGANGKATAN"
i_lbl_working_age(0) = "WORKING AGE (Y)*": i_lbl_working_age(1) = "LAMA BEKERJA"
i_lbl_title(0) = "TITLE": i_lbl_title(1) = "JABATAN"
i_lbl_description(0) = "DESCRIPTION": i_lbl_description(1) = "KETERANGAN"
i_lbl_shiftable(0) = "SHIFTABLE": i_lbl_shiftable(1) = "KERJA SHIFT"
i_lbl_reason_end_working(0) = "END WORKING": i_lbl_reason_end_working(1) = "TERAKHIR BEKERJA"
i_lbl_reason_description(0) = "REASON": i_lbl_reason_description(1) = "ALASAN"

'If mdi_absensi.mnu_stg_language_english.Checked Then
'    i_lang = 0
'ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    i_lang = 1
'End If

lbl_company = i_lbl_company(i_lang)
lbl_employee_code = i_lbl_employee_code(i_lang)
lbl_employee_name = i_lbl_employee_name(i_lang)
lbl_employee_nick_name = i_lbl_employee_nick_name(i_lang)
lbl_department = i_lbl_department(i_lang)
lbl_division = i_lbl_division(i_lang)
fra_last_employment = i_fra_last_employment(i_lang)
lbl_last_employment_name = i_lbl_last_employment_name(i_lang)
lbl_last_employment_date = i_lbl_last_employment_date(i_lang)
lbl_last_employment_title = i_lbl_last_employment_title(i_lang)
lbl_date_of_birth = i_lbl_date_of_birth(i_lang)
lbl_age = i_lbl_age(i_lang)
lbl_place_of_birth = i_lbl_place_of_birth(i_lang)
lbl_sex = i_lbl_sex(i_lang)
lbl_religion = i_lbl_religion(i_lang)
lbl_marital_status = i_lbl_marital_status(i_lang)
lbl_address = i_lbl_address(i_lang)
lbl_phone_number = i_lbl_phone_number(i_lang)
lbl_bank_account = i_lbl_bank_account(i_lang)
fra_last_education = i_fra_last_education(i_lang)
lbl_last_education_level = i_lbl_last_education_level(i_lang)
lbl_last_education_name = i_lbl_last_education_name(i_lang)
lbl_last_education_pass = i_lbl_last_education_pass(i_lang)
lbl_start_working = i_lbl_start_working(i_lang)
lbl_appointment = i_lbl_appointment(i_lang)
lbl_working_age = i_lbl_working_age(i_lang)
lbl_title = i_lbl_title(i_lang)
lbl_description = i_lbl_description(i_lang)
lbl_shiftable = i_lbl_shiftable(i_lang)
lbl_reason_end_working = i_lbl_reason_end_working(i_lang)
lbl_reason_description = i_lbl_reason_description(i_lang)
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select

End Sub

Private Sub txtTelp1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtTelp2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Function get_inc_kode() As String
Dim rs As New ADODB.Recordset
Dim str_inc_kode As String

rs.Open "select max(kode_supplier) as curr_kode from m_supplier", CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    If IsNull(rs.Fields("curr_kode").Value) = True Then
        str_inc_kode = "00001"
    Else
        str_inc_kode = rs.Fields("curr_kode").Value
        str_inc_kode = Right("0000" & Trim(str(CLng(str_inc_kode) + 1)), 5)
    End If
End If

get_inc_kode = str_inc_kode
End Function


Private Sub txtTelp3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub clear_filter()
For Each Col In TDBGrid1.Columns
    Col.FilterText = ""
Next Col
Adodc1.Recordset.Filter = adFilterNone
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

Private Sub opt_active_Click()
Dim strsql As String
Dim v_flag_active As Integer

strsql = "SELECT flag_active FROM m_employee WHERE employee_code = '" & txt_employee_code & "'"
rsBound.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    v_flag_active = rsBound!flag_active
    
If rsBound.RecordCount > 0 Then
    If opt_active Then
        If v_flag_active = "2" Then
            fra_not_active.Visible = False
            lbl_end_mc.Visible = True
            dt_end_mc.Visible = True
            dt_end_mc.Value = Now()
        Else
            fra_not_active.Visible = False
            lbl_end_mc.Visible = False
            dt_end_mc.Visible = False
'            DTPicker_end_working.Value = "NULL"
'            dt_end_mc.Value = "NULL"
        End If
    End If
End If
rsBound.Close
End Sub

Private Sub opt_not_active_Click()
If opt_not_active Then
    fra_not_active.Visible = True
    DTPicker_end_working.Value = Now
    lbl_reason_end_working.Caption = "End Working"
    txt_end_working_reason.Text = ""
    
    lbl_end_mc.Visible = False
    dt_end_mc.Visible = False
End If
End Sub

Private Sub opt_mc_Click()
If opt_mc Then
    fra_not_active.Visible = True
    DTPicker_end_working.Value = Now
    lbl_reason_end_working.Caption = "Start MC"
    txt_end_working_reason.Text = ""
    
    lbl_end_mc.Visible = False
    dt_end_mc.Visible = False
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

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
    
    lbl_employee.Visible = True
    Call load_data_employee
    Call load_data_department
    Call load_data_group_shift
    Call load_count_employee
    
End If
End Sub

Private Sub TDBCombo_company1_ItemChange()
If TDBCombo_company1.ApproxCount > 0 Then
    TDBCombo_company1.Text = TDBCombo_company1.Columns("company_code").Value
    txt_company_name1 = TDBCombo_company1.Columns("company_name").Value
    
    Call load_data_department
    Call load_data_group_shift
End If
End Sub

Private Sub TDBCombo_department_ItemChange()
If TDBCombo_department.ApproxCount > 0 Then
    TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
    txt_department_name = TDBCombo_department.Columns("department_name").Value
    If int_mode = 1 Or int_mode = 2 Then _
        Call load_data_division(TDBCombo_department.Columns("department_code").Value)
End If
End Sub

Private Sub TDBCombo_division_ItemChange()
If TDBCombo_division.ApproxCount > 0 Then
    TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
    txt_division_name = TDBCombo_division.Columns("division_name").Value
End If
End Sub

Private Sub TDBCombo_title_ItemChange()
If TDBCombo_title.ApproxCount > 0 Then
    TDBCombo_title.Text = TDBCombo_title.Columns("title_name").Value
    txt_title_name = TDBCombo_title.Columns("title_code").Value
End If
End Sub

Private Sub TDBCombo_group_shift_ItemChange()
If TDBCombo_group_shift.ApproxCount > 0 Then
    TDBCombo_group_shift.Text = TDBCombo_group_shift.Columns("group_code").Value
    txt_group_shift_name = TDBCombo_group_shift.Columns("group_name").Value
End If
End Sub

Private Sub TDBGrid1_FilterChange()
On Error GoTo ErrHandler

Dim i As Integer

Set Cols = TDBGrid1.Columns
i = TDBGrid1.Col
TDBGrid1.HoldFields

Adodc1.Recordset.Filter = getFilter()
TDBGrid1.Col = i
TDBGrid1.EditActive = True

TDBGrid1.SelStart = Len(TDBGrid1.Columns(i).FilterText)
If TDBGrid1.ApproxCount < 1 Then
    Call clear_filter
    TDBGrid1.Col = i
End If

Exit Sub
ErrHandler:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Public Sub load_data_employee()

active = IIf(optActive.Value, "<> 0", "= 0")
access = IIf(LOGIN_LEVEL = 100, "", "AND (a.managerial_access = 0 OR a.managerial_access IS NULL)")

Adodc1.RecordSource = "SELECT  a.employee_code,a.employee_name,a.employee_nick_name,a.division_code," _
                & "d.division_name,a.department_code,c.department_name,a.company_code," _
                & "b.company_name,a.date_of_birth,a.place_of_birth,a.sex,a.religion," _
                & "a.marital_status,a.number_of_children,a.address,a.email,a.npwp," _
                & "a.phone_number,a.bank_account,a.last_education_code,a.last_education_code_other," _
                & "a.last_education_name,a.last_education_pass,a.last_employment_name," _
                & "a.last_employment_date,a.last_employment_title,a.start_working," _
                & "a.date_of_appointment,a.title_code,e.title_name,a.flag_shiftable," _
                & "a.flag_active,a.description,a.end_working,a.reason,a.flag_pic," _
                & "a.pic,a.access_code,a.managerial_level,a.fathers_name,a.mothers_name," _
                & "a.child_number_from , a.child_number, a.start_mc, a.end_mc, a.branches, a.managerial_access, a.group_shift,a.npwp_address, f.group_name " _
            & "FROM m_employee a LEFT JOIN m_company b ON a.company_code = b.company_code " _
            & "LEFT JOIN m_department c ON a.department_code = c.department_code AND c.company_code = b.company_code " _
            & "LEFT JOIN m_division d ON a.division_code = d.division_code AND d.department_code = c.department_code AND d.company_code = b.company_code " _
            & "LEFT JOIN m_title e ON a.title_code = e.title_code " _
            & "LEFT JOIN m_group_shift f ON a.group_shift = f.group_code and a.company_code = f.company_code " _
            & "where a.company_code = '" _
            & TDBCombo_company.Text & "' " _
            & "AND a.flag_active " & active & " " & access & " order by a.employee_code"
            
'Adodc1.RecordSource = "select * from m_employee where company_code = '" _
'& TDBCombo_company.Columns("company_code").Value & "' order by employee_code"
Adodc1.Refresh

cmdEdit.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
cmdDelete.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
CmdNew.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)
Frame3.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)

TDBGrid1.DataSource = Adodc1
End Sub

Public Sub load_count_employee()
Dim strsql As String

active = IIf(optActive.Value, "<> 0", "= 0")
access = IIf(LOGIN_LEVEL = 100, "", "AND (managerial_access = 0 OR managerial_access IS NULL)")

If rsBound.State = 1 Then rsBound.Close
strsql = "Select Count(employee_code) jml_emp From m_employee " _
        & "WHERE company_code = '" & TDBCombo_company.Text & "' AND flag_active " & active & " " & access & ""
rsBound.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If Not rsBound.EOF Then
    lbl_employee.Caption = "Total Employee = " & rsBound!jml_emp
End If
rsBound.Close

End Sub

Public Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub load_data_department()
'TDBCombo_department.Text = "": txt_department_name = ""

If TDBCombo_company1.Visible = True Then
    Adodc_department.RecordSource = "select * from m_department where company_code='" _
    & TDBCombo_company1.Columns("company_code").Value & "' order by department_code"
    Adodc_department.Refresh
Else
    Adodc_department.RecordSource = "select * from m_department where company_code='" _
    & TDBCombo_company.Columns("company_code").Value & "' order by department_code"
    Adodc_department.Refresh
End If

TDBCombo_department.RowSource = Adodc_department
End Sub

Private Sub load_data_division(ByVal str_department_code As String)
'TDBCombo_division.Text = "": txt_division_name = ""

If TDBCombo_company1.Visible = True Then
    Adodc_division.RecordSource = "select * from m_division where company_code = '" & TDBCombo_company1.Text & "' AND department_code='" _
    & str_department_code & "' order by division_code"
Else
    Adodc_division.RecordSource = "select * from m_division where company_code = '" & TDBCombo_company.Text & "' AND department_code='" _
    & str_department_code & "' order by division_code"
End If
Adodc_division.Refresh

TDBCombo_division.RowSource = Adodc_division
End Sub

Private Sub load_data_title()
Adodc_title.RecordSource = "select * from m_title order by title_code"
Adodc_title.Refresh

TDBCombo_title.RowSource = Adodc_title
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBGrid1.Columns(ColIndex).Caption = "BIRTH DATE" Or _
TDBGrid1.Columns(ColIndex).Caption = "START WORKING" Or _
TDBGrid1.Columns(ColIndex).Caption = "END WORKING" Then
    Value = Format(Value, "yyyy-mm-dd")
End If
End Sub

Private Sub Timer1_Timer()
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Private Sub load_data_group_shift()
'TDBCombo_department.Text = "": txt_department_name = ""

If TDBCombo_company1.Visible = True Then
    Adodc_group_shift.RecordSource = "select * from m_group_shift where company_code='" _
    & TDBCombo_company1.Columns("company_code").Value & "' order by group_code"
    Adodc_group_shift.Refresh
Else
    Adodc_group_shift.RecordSource = "select * from m_group_shift where company_code='" _
    & TDBCombo_company.Columns("company_code").Value & "' order by group_code"
    Adodc_group_shift.Refresh
End If

TDBCombo_group_shift.RowSource = Adodc_group_shift
End Sub

Private Sub txt_address_LostFocus()
If txt_npwp_address = "" Then
    txt_npwp_address.Text = txt_address.Text
End If
End Sub

Private Sub txt_npwp_address_GotFocus()
    txt_npwp_address.SetFocus
    txt_npwp_address.SelStart = 0
    txt_npwp_address.SelLength = Len(txt_npwp_address.Text)
End Sub
