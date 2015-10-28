VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_rpt_summary_employee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LAPORAN - REKAP KARYAWAN"
   ClientHeight    =   7095
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   10560
   Icon            =   "frm_rpt_summary_employee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3330
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Top             =   750
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   1230
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "PARAMETER"
      TabPicture(0)   =   "frm_rpt_summary_employee.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   2775
         Left            =   840
         TabIndex        =   5
         Top             =   840
         Width           =   8415
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   3720
            TabIndex        =   13
            Top             =   1590
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   46202883
            CurrentDate     =   41234
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
            Left            =   3720
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   1200
            Width           =   2655
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_department 
            Height          =   375
            Left            =   3720
            OleObjectBlob   =   "frm_rpt_summary_employee.frx":05A6
            TabIndex        =   8
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PERIODE PENGGAJIAN"
            Height          =   195
            Left            =   1800
            TabIndex        =   12
            Top             =   1620
            Width           =   1785
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "DEPARTEMEN"
            Height          =   195
            Left            =   1800
            TabIndex        =   6
            Top             =   870
            Width           =   1125
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   5670
      Width           =   10095
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
         Left            =   8550
         Picture         =   "frm_rpt_summary_employee.frx":2567
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
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
         Left            =   7320
         Picture         =   "frm_rpt_summary_employee.frx":2AF1
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   30
         Top             =   150
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1530
      OleObjectBlob   =   "frm_rpt_summary_employee.frx":307B
      TabIndex        =   3
      Top             =   750
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "LAPORAN REKAP KARYAWAN"
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
      Left            =   3720
      TabIndex        =   9
      Top             =   150
      Width           =   4365
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   810
      Width           =   1110
   End
End
Attribute VB_Name = "frm_rpt_summary_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rscompany As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset

Dim v_access As String
Dim v_sex As String
Dim v_age As String
Dim v_working_age As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub rpt_summary_employee()
Dim str_file As String
Dim a As New frm_rpt
Dim dt1 As String
Dim dt2 As String
    
    str_file = "\report\rpt_master_employee.rpt"
    
    dt1 = Format(DTPicker1.Value, "yyyy-MM") & "-01"
    dt2 = Format(DTPicker1.Value, "yyyy-MM") & "-" & getEndDay(Format(DTPicker1.Value, "MM"), Format(DTPicker1.Value, "yyyy"))
    
    v_access = IIf(LOGIN_LEVEL = 100, "", "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND flag_active = 0 order by employee_name")
    
    SQL = "SELECT * FROM (SELECT a.company_code,b.company_name,a.department_code,c.department_name," & _
                   "a.nik,a.employee_name,a.title_code,e.title_name, " & _
                   "IFNULL((SELECT flag_ot FROM m_salary_standard f WHERE f.employee_code = a.employee_code AND f.salary_date <= '" & dt2 & "' ORDER BY f.number DESC LIMIT 1),0) " & _
                "FROM m_employee a JOIN m_company b ON a.company_code = b.company_code " & _
                "JOIN m_department c ON a.department_code = c.department_code AND c.company_code = b.company_code " & _
                "JOIN m_title e ON a.title_code = e.title_code " & _
                "WHERE " & IIf(TDBCombo_department.Text <> "", "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND a.department_code = '" & TDBCombo_department.Text & "'", "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "'") & " " & _
                    "" & access & " " & _
                    "AND date(start_working) < '" & dt2 & "' AND flag_active <> 0 " & _
                "ORDER BY employee_name ASC) aa "
                
    Call a.Show
    
    a.Caption = "REPORT SUMMARY EMPLOYEE"
    Call a.rpt_view(SQL, str_file, Now)
End Sub

Private Sub cmdPrint_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        Call rpt_summary_employee
    End If
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Now
    
    Call load_data_company
    Call load_data_user_access(Me)
    
    Timer1.Enabled = True
End Sub

Private Sub load_data_company()
    If rscompany.State Then rscompany.Close
    SQL = "select * from m_company order by company_code"
    rscompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rscompany
End Sub

Private Sub load_data_department()
    If rsDept.State Then rsDept.Close
    SQL = "select * from m_department where company_code = '" & TDBCombo_company.Text & "' order by department_code"
    rsDept.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_department.RowSource = rsDept
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If
    
    Call load_data_department
End Sub

Private Sub TDBCombo_department_Change()
    If TDBCombo_department.Text = "" Then txt_department_name.Text = ""
End Sub

Private Sub TDBCombo_department_ItemChange()
    If TDBCombo_department.ApproxCount > 0 Then
        TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
        txt_department_name = TDBCombo_department.Columns("department_name").Value
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call set_company_mode1(rscompany, TDBCombo_company, txt_company_name)
End Sub
