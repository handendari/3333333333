VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_rpt_salary_bank 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LAPORAN - DISTRIBUSI BANK"
   ClientHeight    =   7395
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   10560
   Icon            =   "frm_rpt_salary_bank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Text            =   "frm_rpt_salary_bank.frx":058A
      Top             =   870
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Top             =   990
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   1470
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
      TabPicture(0)   =   "frm_rpt_salary_bank.frx":0590
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra_process_monthly"
      Tab(0).Control(1)=   "fra_monthly"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "PERIODE"
      TabPicture(1)   =   "frm_rpt_salary_bank.frx":05AC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fra_process_periode"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra_periode"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fra_monthly 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   5
         Top             =   660
         Width           =   8415
         Begin VB.ComboBox cbo_monthly_company 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":05C8
            Left            =   1800
            List            =   "frm_rpt_salary_bank.frx":05D2
            TabIndex        =   11
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.Frame fra_monthly_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   7
            Top             =   960
            Width           =   4575
            Begin VB.CommandButton cmd_monthly_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1440
               TabIndex        =   10
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
               TabIndex        =   9
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
               TabIndex        =   8
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.ComboBox cbo_monthly_employee 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":05E5
            Left            =   1800
            List            =   "frm_rpt_salary_bank.frx":05EF
            TabIndex        =   6
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_monthly 
            Height          =   300
            Left            =   1800
            TabIndex        =   12
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   95223811
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label13 
            Caption         =   "* yyyy-MM"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   3540
            TabIndex        =   38
            Top             =   1590
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Perusahaan"
            Height          =   195
            Left            =   720
            TabIndex        =   15
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Bulan"
            Height          =   195
            Left            =   720
            TabIndex        =   14
            Top             =   1560
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Karyawan"
            Height          =   195
            Left            =   720
            TabIndex        =   13
            Top             =   1200
            Width           =   705
         End
      End
      Begin VB.Frame fra_process_monthly 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   29
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
            TabIndex        =   30
            Top             =   1080
            Width           =   2730
         End
      End
      Begin VB.Frame fra_periode 
         Height          =   2925
         Left            =   840
         TabIndex        =   16
         Top             =   660
         Width           =   8415
         Begin VB.ComboBox cbo_periode_to 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":0600
            Left            =   3600
            List            =   "frm_rpt_salary_bank.frx":060A
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "..."
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txt_periode_employee_code 
            Height          =   285
            Left            =   4590
            TabIndex        =   37
            Text            =   "Text2"
            Top             =   180
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Frame fra_periode_department 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   3600
            TabIndex        =   34
            Top             =   420
            Width           =   4695
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
               Height          =   285
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               MultiLine       =   -1  'True
               TabIndex        =   35
               Top             =   90
               Width           =   2415
            End
            Begin TrueOleDBList60.TDBCombo TDBCombo_division 
               Height          =   375
               Left            =   0
               OleObjectBlob   =   "frm_rpt_salary_bank.frx":0618
               TabIndex        =   36
               Top             =   90
               Width           =   1815
            End
            Begin MSAdodcLib.Adodc Adodc_division 
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
         Begin VB.ComboBox cbo_periode_division 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":257F
            Left            =   1800
            List            =   "frm_rpt_salary_bank.frx":2589
            TabIndex        =   33
            Text            =   "..."
            Top             =   510
            Width           =   1695
         End
         Begin VB.ComboBox cbo_bank_name 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":259C
            Left            =   1800
            List            =   "frm_rpt_salary_bank.frx":259E
            TabIndex        =   31
            Text            =   "..."
            Top             =   2070
            Width           =   1695
         End
         Begin VB.ComboBox cbo_periode_company 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":25A0
            Left            =   3570
            List            =   "frm_rpt_salary_bank.frx":25AA
            TabIndex        =   23
            Text            =   "..."
            Top             =   2070
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "DAY COUNT"
            Height          =   495
            Left            =   0
            TabIndex        =   22
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cbo_periode_employee 
            Height          =   315
            ItemData        =   "frm_rpt_salary_bank.frx":25BD
            Left            =   1800
            List            =   "frm_rpt_salary_bank.frx":25C7
            TabIndex        =   21
            Text            =   "..."
            Top             =   900
            Width           =   1695
         End
         Begin VB.Frame fra_periode_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   585
            Left            =   3600
            TabIndex        =   17
            Top             =   660
            Width           =   4575
            Begin VB.TextBox txt_periode_nik 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   0
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   20
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
               TabIndex        =   19
               Top             =   240
               Width           =   2415
            End
            Begin VB.CommandButton cmd_periode_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1440
               TabIndex        =   18
               Top             =   240
               Width           =   375
            End
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_from 
            Height          =   300
            Left            =   1800
            TabIndex        =   40
            Top             =   1680
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   95223811
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_to 
            Height          =   300
            Left            =   5040
            TabIndex        =   41
            Top             =   1680
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   95223811
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   1800
            TabIndex        =   42
            Top             =   1290
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   95223811
            CurrentDate     =   39278
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BULAN"
            Height          =   195
            Left            =   1170
            TabIndex        =   44
            Top             =   1290
            Width           =   540
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PERIODE"
            Height          =   195
            Left            =   990
            TabIndex        =   43
            Top             =   1710
            Width           =   720
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BANK"
            Height          =   195
            Left            =   1290
            TabIndex        =   32
            Top             =   2130
            Width           =   435
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DIVISI"
            Height          =   195
            Left            =   1230
            TabIndex        =   25
            Top             =   570
            Width           =   465
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KARYAWAN"
            Height          =   195
            Left            =   780
            TabIndex        =   24
            Top             =   900
            Width           =   930
         End
      End
      Begin VB.Frame fra_process_periode 
         Height          =   2655
         Left            =   840
         TabIndex        =   27
         Top             =   660
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
            TabIndex        =   28
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
      Top             =   5850
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   840
         Top             =   360
      End
      Begin prj_panji.vbButton cmdExit 
         Height          =   705
         Left            =   8490
         TabIndex        =   45
         Top             =   300
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Keluar"
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
         MICON           =   "frm_rpt_salary_bank.frx":25D8
         PICN            =   "frm_rpt_salary_bank.frx":25F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_print_bank 
         Height          =   705
         Left            =   6330
         TabIndex        =   46
         Top             =   300
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Cetak"
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
         MICON           =   "frm_rpt_salary_bank.frx":3686
         PICN            =   "frm_rpt_salary_bank.frx":36A2
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1470
      OleObjectBlob   =   "frm_rpt_salary_bank.frx":4734
      TabIndex        =   3
      Top             =   990
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1350
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
      Left            =   9000
      Top             =   630
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
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "LAPORAN DISTRIBUSI BANK"
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
      Left            =   240
      TabIndex        =   47
      Top             =   180
      Width           =   4155
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1050
      Width           =   1110
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_rpt_salary_bank.frx":669A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12690
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

Private Sub cbo_periode_division_Click()
    TDBCombo_division.Text = ""
    txt_division_name.Text = ""
        
    If cbo_periode_division.ListIndex = 0 Then
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

Private Sub cmd_periode_browse_employee_Click()
    frm_lookup_mst_employee.public_int_mode = 79
    frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
    frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_print_bank_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If

    If SSTab1.Tab = 1 Then
        Call rpt_bank
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Function check_valid_periode() As Boolean
    check_valid_periode = True
    
    'validate employee
    If cbo_periode_employee.ListIndex = 1 And Trim(txt_periode_employee_code) = "" Then
        MsgBox "Karyawan Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
        cmd_periode_browse_employee.SetFocus
        check_valid_periode = False
        Exit Function
    End If
End Function

Private Sub rpt_bank()
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer

    int_process = vbNo
    
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
    
    str_sql = "call spr_salary_hardys_bank('" & d1 & "','" & d2 & "'," _
            & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_division.Text & "'," _
            & int_flag_employee & ",'" & str_employee_code & "','" & cbo_bank_name.Text & "'," _
            & cbo_periode_division.ListIndex & ",'" & LOGIN_LEVEL & "', '" & LOGIN_CODE & "')"
    
    Text1 = str_sql
    
    Call a.Show
    a.Caption = "SALARY BANK REPORT"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub Form_Load()
    Adodc_company.ConnectionString = strConn
    Adodc_division.ConnectionString = strConn

    Call load_data_company
    'Call load_data_setting_mail
    Call load_data_bank
    Call load_data_user_access(Me)

    DTPicker_periode_from.Value = Now
    DTPicker_periode_to.Value = Now
    DTPicker_monthly.Value = Now
    DTPicker1.Value = Now
    
    Call getPeriodeProses
    
    cbo_periode_to.ListIndex = 0
    cbo_periode_company.ListIndex = 1
    cbo_periode_employee.ListIndex = 0
    
    timer1.Enabled = True
    SSTab1.TabVisible(0) = False
    cbo_bank_name.ListIndex = 0
    cbo_periode_division.ListIndex = 0
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
    
    Call load_data_division
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode_adodc(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Private Sub set_data_department(ByVal str_code As String)
On Error Resume Next

    Adodc_division.Recordset.MoveFirst
    Adodc_division.Recordset.Find ("department_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (Adodc_division.Recordset.EOF = True Or Adodc_division.Recordset.BOF = True) Then
        TDBCombo_division.Bookmark = Adodc_division.Recordset.AbsolutePosition
        Call tdbcombo_division_itemChange
    Else
        TDBCombo_division.Text = ""
    End If
End Sub

Private Sub load_data_division()
TDBCombo_division.Text = "": txt_division_name = ""

    Adodc_division.RecordSource = "select division_code, division_name from m_division where company_code='" _
    & TDBCombo_company.Columns("company_code").Value & "' order by division_code"
    Adodc_division.Refresh
    
    TDBCombo_division.RowSource = Adodc_division
End Sub

Private Sub tdbcombo_division_itemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
        txt_division_name = TDBCombo_division.Columns("division_name").Value
    End If
End Sub

Private Sub DTPicker1_Change()
    Call getPeriodeProses
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    Call getPeriodeProses
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

Private Sub getPeriodeProses()
Dim strsql As String
Dim rsbulan As New ADODB.Recordset

    strsql = "select a.date_from,a.date_to from h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
            & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " _
            & "AND left(`month`,7) = '" & Format(DTPicker1.Value, "yyyy-MM") & "'"
    rsbulan.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsbulan.RecordCount > 0 Then
        DTPicker_periode_from.Value = rsbulan!date_from
        DTPicker_periode_to.Value = rsbulan!date_to
        cbo_periode_to.ListIndex = 1
    End If
End Sub

