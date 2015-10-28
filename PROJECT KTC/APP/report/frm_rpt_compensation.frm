VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_rpt_compensation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "REPORT COMPENSATION"
   ClientHeight    =   6855
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   10560
   Icon            =   "frm_rpt_compensation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Text            =   "frm_rpt_compensation.frx":000C
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3300
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   10
      Top             =   240
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "MONTHLY"
      TabPicture(0)   =   "frm_rpt_compensation.frx":0012
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra_process_monthly"
      Tab(0).Control(1)=   "fra_monthly"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "PERIODE"
      TabPicture(1)   =   "frm_rpt_compensation.frx":002E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fra_process_periode"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra_periode"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fra_monthly 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   12
         Top             =   960
         Width           =   8415
         Begin VB.TextBox txt_title 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1920
            Width           =   3135
         End
         Begin VB.ComboBox cbo_monthly_company 
            Height          =   315
            ItemData        =   "frm_rpt_compensation.frx":004A
            Left            =   1800
            List            =   "frm_rpt_compensation.frx":0054
            TabIndex        =   2
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.Frame fra_monthly_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   13
            Top             =   960
            Width           =   4575
            Begin VB.CommandButton cmd_monthly_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1440
               TabIndex        =   4
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
               TabIndex        =   15
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
               TabIndex        =   14
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.ComboBox cbo_monthly_employee 
            Height          =   315
            ItemData        =   "frm_rpt_compensation.frx":0067
            Left            =   1800
            List            =   "frm_rpt_compensation.frx":0071
            TabIndex        =   3
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_monthly 
            Height          =   300
            Left            =   1800
            TabIndex        =   5
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   126877699
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "TITLE"
            Height          =   195
            Left            =   720
            TabIndex        =   38
            Top             =   1950
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "BRANCH OFFICE"
            Height          =   195
            Left            =   330
            TabIndex        =   18
            Top             =   870
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "MONTH"
            Height          =   195
            Left            =   720
            TabIndex        =   17
            Top             =   1560
            Width           =   600
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   720
            TabIndex        =   16
            Top             =   1200
            Width           =   870
         End
      End
      Begin VB.Frame fra_process_monthly 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   36
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
            TabIndex        =   37
            Top             =   1080
            Width           =   2730
         End
      End
      Begin VB.Frame fra_periode 
         Height          =   2655
         Left            =   840
         TabIndex        =   19
         Top             =   960
         Width           =   8415
         Begin VB.TextBox txt_title1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   39
            Top             =   1920
            Width           =   3135
         End
         Begin VB.ComboBox cbo_periode_to 
            Height          =   315
            ItemData        =   "frm_rpt_compensation.frx":0082
            Left            =   3600
            List            =   "frm_rpt_compensation.frx":008C
            TabIndex        =   27
            Text            =   "..."
            Top             =   1560
            Width           =   1335
         End
         Begin VB.ComboBox cbo_periode_company 
            Height          =   315
            ItemData        =   "frm_rpt_compensation.frx":009A
            Left            =   1800
            List            =   "frm_rpt_compensation.frx":00A4
            TabIndex        =   26
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "DAY COUNT"
            Height          =   495
            Left            =   0
            TabIndex        =   25
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cbo_periode_employee 
            Height          =   315
            ItemData        =   "frm_rpt_compensation.frx":00B7
            Left            =   1800
            List            =   "frm_rpt_compensation.frx":00C1
            TabIndex        =   24
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Frame fra_periode_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   20
            Top             =   960
            Width           =   4575
            Begin VB.TextBox txt_periode_employee_code 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   0
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   23
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
               TabIndex        =   22
               Top             =   240
               Width           =   2415
            End
            Begin VB.CommandButton cmd_periode_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1440
               TabIndex        =   21
               Top             =   240
               Width           =   375
            End
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_from 
            Height          =   300
            Left            =   1800
            TabIndex        =   28
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   126877699
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_to 
            Height          =   300
            Left            =   5040
            TabIndex        =   29
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   126877699
            CurrentDate     =   39278
         End
         Begin VB.Label TITLE 
            Caption         =   "TITLE"
            Height          =   285
            Left            =   720
            TabIndex        =   40
            Top             =   1920
            Width           =   825
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PERIODE"
            Height          =   195
            Left            =   720
            TabIndex        =   32
            Top             =   1560
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "COMPANY"
            Height          =   195
            Left            =   720
            TabIndex        =   31
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   720
            TabIndex        =   30
            Top             =   1200
            Width           =   870
         End
      End
      Begin VB.Frame fra_process_periode 
         Height          =   2655
         Left            =   840
         TabIndex        =   34
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
            TabIndex        =   35
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
      Top             =   5130
      Width           =   10095
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
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
         Left            =   6660
         Picture         =   "frm_rpt_compensation.frx":00D2
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   8850
         Picture         =   "frm_rpt_compensation.frx":065C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1560
      OleObjectBlob   =   "frm_rpt_compensation.frx":0BE6
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "BRANCH OFFICE"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   300
      Width           =   1275
   End
End
Attribute VB_Name = "frm_rpt_compensation"
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
    txt_monthly_employee_code.Text = ""
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

Private Sub cbo_periode_employee_Click()
If cbo_periode_employee.ListIndex = 0 Then
    fra_periode_employee.Visible = False
    txt_periode_employee_code.Text = ""
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

Private Sub cmd_print_slip_meal_Click()
Dim strsql As String
Dim pTgl1 As String, pTgl2 As String, str_file As String, kdkaryawan As String
Dim a As New frm_rpt
    
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Branch Office selected!", vbInformation, headerMSG
        Exit Sub
    End If

    ' str_param1 = Format(DTPicker1.Value, "dd-MMM-yyyy")
    
    If SSTab1.Tab = 0 Then
        pTgl1 = Format(DTPicker_monthly.Value, "yyyy-MM-01")
        pTgl2 = Format(DTPicker_monthly.Value, "yyyy-MM-31")
    Else
        pTgl1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
        pTgl2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
    End If
    
    str_file = "\report\rpt_meal.rpt"
    
    If cbo_monthly_employee.ListIndex = 0 Then
        strsql = "SELECT a.employee_code, a.employee_name, a.title_name, a.start_working, " & _
                "'" & Format(DTPicker_monthly.Value, "yyyy-MM-dd") & "', b.salary, " & _
                "CASE WHEN a.marital_status = 0 THEN 'S' ELSE CONCAT('M/',CAST(IFNULL(a.number_of_children,0) AS CHAR)) END marital_status, " & _
                "fn_jmlHariKerja(a.employee_code,'" & pTgl1 & "','" & pTgl2 & "') jmlHariMasuk, b.uangmakan " & _
                "FROM m_employee a JOIN m_salary b ON a.employee_code = b.employee_code " & _
                "JOIN h_attendance c ON b.employee_code = c.employee_code " & _
                "WHERE a.company_code = '" & TDBCombo_company.Text & "' AND " & _
                "MONTH(c.att_date) = MONTH('" & pTgl1 & "') AND " & _
                "YEAR(c.att_date) = YEAR('" & pTgl1 & "') " & _
                "GROUP BY a.employee_code"
    Else
        strsql = "SELECT a.employee_code, a.employee_name, a.title_name, a.start_working, " & _
                "'" & Format(DTPicker_monthly.Value, "yyyy-MM-dd") & "', b.salary, " & _
                "CASE WHEN a.marital_status = 0 THEN 'S' ELSE CONCAT('M/',CAST(IFNULL(a.number_of_children,0) AS CHAR)) END marital_status, " & _
                "fn_jmlHariKerja(a.employee_code,'" & pTgl1 & "','" & pTgl2 & "') jmlHariMasuk, b.uangmakan " & _
                "FROM m_employee a JOIN m_salary b ON a.employee_code = b.employee_code " & _
                "JOIN h_attendance c ON b.employee_code = c.employee_code " & _
                "WHERE a.employee_code = '" & txt_monthly_employee_code.Text & "' AND " & _
                "a.company_code = '" & TDBCombo_company.Text & "' AND " & _
                "MONTH(c.att_date) = MONTH('" & pTgl1 & "') AND " & _
                "YEAR(c.att_date) = YEAR('" & pTgl1 & "') " & _
                "GROUP BY a.employee_code"
    End If
            
    Call a.Show
    a.Caption = "SLIP MEAL ALLOANCE"
    'str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
    Call a.rpt_view(strsql, str_file, pTgl1)
End Sub

'Private Sub cmd_send_mail_Click()
'If check_validate_tdbcombo(TDBCombo_company) = False Then
'    MsgBox "No Company selected!", vbInformation, headerMSG
'    Exit Sub
'End If
'
'int_sent_false = 0
'int_sent_true = 0
'
'If SSTab1.Tab = 1 Then
'    If check_valid_periode Then
'        Call send_mail_periode(1)
'    End If
'ElseIf SSTab1.Tab = 0 Then
'    If check_valid_monthly Then
'        Call send_mail_monthly(1)
'    End If
'
'End If
'End Sub

'Private Sub send_mail_monthly(ByVal j As Integer)
'Dim str_sql, str_param_periode, str_file, str1, str_file_out As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim dt1, dt2 As Date
'Dim d1, d2 As String
'Dim a As New frm_rpt
'
'If j = 0 Then
'    str_file = "\report\rpt_slip_gajiA.rpt"
'ElseIf j = 1 Then
'    str_file = "\report\rpt_slip_gajiA.rpt"
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
'    pub_employee_name = rs_employee!employee_name
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
'
'Private Sub send_mail_periode(ByVal j As Integer)
'Dim str_sql, str_param_periode, str_file, str1, str_file_out As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim a As New frm_rpt
'
'If j = 0 Then
'    str_file = "\report\rpt_slip_gajiA.rpt"
'ElseIf j = 1 Then
'    str_file = "\report\rpt_slip_gajiA.rpt"
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
'    pub_employee_name = rs_employee!employee_name
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

Private Sub show_msg()
MsgBox "There are " & int_sent_true & " mail are sent successfully!" & vbCrLf _
    & int_sent_false & " are fails!" & vbCrLf _
    & "For more detail info, let see the 'log mail'!", vbInformation, headerMSG
End Sub

Private Sub rpt_auto_pdf(ByVal sql_proc As String, ByVal sql_proc2 As String, ByVal sql_proc3 As String, ByVal rpt_file As String, _
ByVal str_file_out As String, ByVal str_param As String)
Dim CrApp As New CRAXDRT.Application
Dim CrRep As New CRAXDRT.Report
Dim AdoRs As New ADODB.Recordset
Dim AdoRs2 As New ADODB.Recordset
Dim Adors3 As New ADODB.Recordset

Dim CrDatabase As CRAXDRT.Database
Dim CrDatabaseTables As CRAXDRT.DatabaseTables
Dim CrDatabaseTable As CRAXDRT.DatabaseTable
Dim CrSections As CRAXDRT.Sections
Dim CrSection As CRAXDRT.Section
Dim CrReportObjs As CRAXDRT.ReportObjects
Dim CrSubreportObj As CRAXDRT.SubreportObject
Dim CrSubreport As CRAXDRT.Report

AdoRs.Open sql_proc, CnG, adOpenDynamic, adLockBatchOptimistic
AdoRs2.Open sql_proc2, CnG, adOpenDynamic, adLockBatchOptimistic
Adors3.Open sql_proc3, CnG, adOpenDynamic, adLockBatchOptimistic

Set CrRep = CrApp.OpenReport(App.Path & rpt_file)
CrRep.DiscardSavedData
CrRep.Database.Tables(1).SetDataSource AdoRs, 3
CrRep.ParameterFields.GetItemByName("p_param1").AddCurrentValue str_param
'---------------------
Dim x As Integer
Dim Y As Integer
Dim tabel As String

    Set CrSections = CrRep.Sections
    For x = 1 To CrSections.Count
        Set CrSection = CrSections.item(x)
        Set CrReportObjs = CrSection.ReportObjects
        For Y = 1 To CrReportObjs.Count
            If CrReportObjs.item(Y).Kind = crSubreportObject Then
            
                Set CrSubreportObj = CrReportObjs.item(Y)
                Set CrSubreport = CrSubreportObj.OpenSubreport
                Set CrDatabase = CrSubreport.Database
    
                Set CrDatabaseTables = CrDatabase.Tables
                Set CrDatabaseTable = CrDatabaseTables.item(1)
                tabel = CrDatabaseTables.item(1).name
                If tabel = "ttx_slip_gaji_detail2A_ttx" Then
                    CrDatabaseTable.SetDataSource Adors3, 3
                Else
                   CrDatabaseTable.SetDataSource AdoRs2, 3
                End If
                
                
                'MsgBox "Table : " & CrDatabaseTable.Name & " in subreport " _
                    & UCase(CrSubreportObj.SubreportName) & " in location " _
                    & CrDatabaseTable.Location
                
            End If
        Next
           
    Next
    
'    crv.ReportSource = CrRep
'    crv.viewReport
'    Me.MousePointer = vbNormal

'---
CrRep.ExportOptions.DestinationType = crEDTDiskFile
CrRep.ExportOptions.FormatType = crEFTPortableDocFormat
CrRep.ExportOptions.DiskFileName = str_file_out
CrRep.Export False
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

Private Sub Command2_Click()
    Dim strsql As String
    Dim pTgl1 As String, pTgl2 As String, str_file As String, kdkaryawan As String
    Dim a As New frm_rpt
    
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Branch Office selected!", vbInformation, headerMSG
        Exit Sub
    End If

    If SSTab1.Tab = 0 Then
        pTgl1 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-01"
        pTgl2 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-" & getEndDay(DTPicker_monthly.Month)
        kdkaryawan = txt_monthly_employee_code.Text
    Else
        pTgl1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
        pTgl2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
        kdkaryawan = txt_periode_employee_code.Text
    End If
    
    str_file = "\report\rpt_ot_record.rpt"
    
    strsql = "call sp_rekap_overtime('" & kdkaryawan & "','" & pTgl1 & "','" & pTgl2 & "','" & TDBCombo_company.Text & "')"
    Call a.Show
    a.Caption = "SUMMARY OVERTIME REPORT"
    'str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
     
    Call a.rpt_view(strsql, str_file, pTgl1)
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

Private Sub rpt_periode(ByVal j As Integer)
Dim str_sql, str_param_periode, str_file, str1, str2  As String
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
    str_sql = "call spr_salary_ebiz_sum('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
ElseIf cbo_periode_to.ListIndex = 1 Then
    str_sql = "call spr_salary_ebiz_sum('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
        & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
End If
'--
ElseIf j = 1 Then
'--
If cbo_periode_to.ListIndex = 0 Then
    str_sql = "call spr_salary_ebiz_sli('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
ElseIf cbo_periode_to.ListIndex = 1 Then
    str_sql = "call spr_salary_ebiz_sli('" & d1 & "','" & d2 & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
        & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
End If
'--
End If

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
    Dim strsql As String
    Dim pTgl1 As String, pTgl2 As String, str_file As String, kdkaryawan As String
    Dim a As New frm_rpt
    
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Branch Office selected!", vbInformation, headerMSG
        Exit Sub
    End If
  
    If SSTab1.Tab = 0 Then
        pTgl1 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-01"
        pTgl2 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-" & getEndDay(DTPicker_monthly.Month)
        kdkaryawan = txt_monthly_employee_code.Text
    Else
        pTgl1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
        pTgl2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
        kdkaryawan = txt_periode_employee_code.Text
    End If
    
    str_file = "\report\rpt_summary_payroll2.rpt"
    
    strsql = "call sp_summary_payroll('" & kdkaryawan & "','" & pTgl1 & "','" & pTgl2 & "','" & TDBCombo_company.Text & "')"
    Call a.Show
    a.Caption = "SUMMARY SALARY REPORT"
    'str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
     
    Call a.rpt_view(strsql, str_file, pTgl1)
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
    Dim strsql As String
    Dim pTgl1 As String, pTgl2 As String, str_file As String, kdkaryawan As String
    Dim a As New frm_rpt
    Dim judul As String
    Dim str_param_periode As String
    
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Branch Office selected!", vbInformation, headerMSG
        Exit Sub
    End If

    If SSTab1.Tab = 0 Then
        pTgl1 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-01"
        pTgl2 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-" & getEndDay(DTPicker_monthly.Month)
        kdkaryawan = txt_monthly_employee_code.Text
        judul = txt_title.Text
        str_param_periode = "MONTH : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
    Else
        pTgl1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
        pTgl2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
        kdkaryawan = txt_periode_employee_code.Text
        judul = txt_title1.Text
        str_param_periode = "PERIODE : " & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & " s/d " & "PERIODE : " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
    End If
    
    str_file = "\report\rpt_kompensasi.rpt"
    
    strsql = "call sp_Kompensasi('" & kdkaryawan & "','" & TDBCombo_company.Text & "','" & judul & "')"
    Call a.Show
    a.Caption = "REPORT COMPENSATION"
    'str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
     
    Call a.rpt_view(strsql, str_file, str_param_periode)
End Sub

Private Sub Command4_Click()
    Dim strsql As String
    Dim pTgl1 As String, pTgl2 As String, str_file As String, kdkaryawan As String
    Dim a As New frm_rpt
    
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Branch Office selected!", vbInformation, headerMSG
        Exit Sub
    End If

    If SSTab1.Tab = 0 Then
        pTgl1 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-01"
        pTgl2 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-" & getEndDay(DTPicker_monthly.Month)
        kdkaryawan = txt_monthly_employee_code.Text
    Else
        pTgl1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
        pTgl2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
        kdkaryawan = txt_periode_employee_code.Text
    End If
    
    str_file = "\report\rpt_site_allowance.rpt"
    
    strsql = "call sp_rekap_siteAllowance('" & kdkaryawan & "','" & pTgl1 & "','" & pTgl2 & "'," _
            & "'" & TDBCombo_company.Text & "')"
    Call a.Show
    a.Caption = "SUMMARY SITE ALLOWANCE REPORT"
    'str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
     
    Call a.rpt_view(strsql, str_file, pTgl1)

End Sub

Private Sub Command5_Click()
    Dim strsql As String
    Dim pTgl1 As String, pTgl2 As String, str_file As String
    Dim kdkaryawan As String
    Dim a As New frm_rpt
    
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Branch Office selected!", vbInformation, headerMSG
        Exit Sub
    End If

    If SSTab1.Tab = 0 Then
        pTgl1 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-01"
        pTgl2 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-" & getEndDay(DTPicker_monthly.Month)
        kdkaryawan = txt_monthly_employee_code.Text
    Else
        pTgl1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
        pTgl2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
        kdkaryawan = txt_periode_employee_code.Text
    End If
    
    str_file = "\report\rpt_terima_potong_lain.rpt"
    
    strsql = "call sp_rekap_terima_potong_lain('" & kdkaryawan & "','" & pTgl1 & "','" & pTgl2 & "','" & TDBCombo_company.Text & "')"
    Call a.Show
    a.Caption = "SUMMARY OTHER INCOME & EXPENSE REPORT"
    'str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
     
    Call a.rpt_view(strsql, str_file, pTgl1)
End Sub

Private Sub Command6_Click()
Dim strsql As String
    Dim pTgl1 As String, pTgl2 As String, str_file As String, kdkaryawan As String
    Dim a As New frm_rpt
    
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Branch Office selected!", vbInformation, headerMSG
        Exit Sub
    End If

    
    If SSTab1.Tab = 0 Then
        pTgl1 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-01"
        pTgl2 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-" & getEndDay(DTPicker_monthly.Month)
        kdkaryawan = txt_monthly_employee_code.Text
    Else
        pTgl1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
        pTgl2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
        kdkaryawan = txt_periode_employee_code.Text
    End If
    
    str_file = "\report\rpt_summary_payroll2.rpt"
    
    strsql = "call sp_summary_payroll('" & kdkaryawan & "','" & pTgl1 & "','" & pTgl2 & "')"
    Call a.Show
    a.Caption = "SUMMARY SALARY REPORT"
    'str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
     
    Call a.rpt_view(strsql, str_file, pTgl1)
End Sub

Private Sub Form_Load()
Adodc_company.ConnectionString = strConn

Call load_data_company
Call load_data_setting_mail

DTPicker_periode_from.Value = Now
DTPicker_periode_to.Value = Now
DTPicker_monthly.Value = Now

cbo_periode_to.ListIndex = 0
cbo_periode_company.ListIndex = 1
cbo_periode_employee.ListIndex = 0

cbo_monthly_company.ListIndex = 1
cbo_monthly_employee.ListIndex = 0

Timer1.Enabled = True
SSTab1.Tab = 0
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
End Sub

Private Sub load_data_setting_mail()
Dim rs1 As New ADODB.Recordset
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


rs1.Open "select * from s_mail where s_number=1", CnG, adOpenStatic, adLockReadOnly
If rs1.RecordCount > 0 Then
    'str_smtp = rs1.Fields("s_smtp").Value
    str_sender_mail = rs1.Fields("s_sender_email").Value
    str_sender_name = rs1.Fields("s_sender_name").Value
Else
    str_smtp = ""
    str_sender_mail = ""
    str_sender_name = ""
    MsgBox "No valid SMTP setting!", vbInformation, headerMSG
    'Unload Me
End If
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
    '.SMTPHost = str_smtp
    .SMTPHost = "mail.solusisentraldata.com"
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
        .Username = "send.ktc@solusisentraldata.com"  ' Optional, default = Null String
        .Password = "sepatukuda2011"                            ' Optional, default = Null String, value is NOT saved
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

Private Sub txt_title_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_title1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
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



