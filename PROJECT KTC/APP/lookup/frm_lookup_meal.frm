VERSION 5.00
Begin VB.Form frm_lookup_meal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PRINT MEAL"
   ClientHeight    =   2925
   ClientLeft      =   -15
   ClientTop       =   345
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_close 
      Caption         =   "&Close"
      Height          =   645
      Left            =   3450
      Picture         =   "frm_lookup_meal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmd_use 
      Caption         =   "&OK"
      Height          =   645
      Left            =   2370
      Picture         =   "frm_lookup_meal.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print Mode"
      Height          =   1905
      Left            =   360
      TabIndex        =   0
      Top             =   150
      Width           =   4065
      Begin VB.Frame frm_summary 
         Caption         =   "Summary Meal"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   3825
         Begin VB.TextBox txt_check 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1230
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   420
            Width           =   2475
         End
         Begin VB.Label Label2 
            Caption         =   "Checked by"
            Height          =   225
            Left            =   150
            TabIndex        =   6
            Top             =   450
            Width           =   945
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Summary Meal"
         Height          =   315
         Index           =   1
         Left            =   1890
         TabIndex        =   4
         Top             =   390
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Slip Meal"
         Height          =   315
         Index           =   0
         Left            =   660
         TabIndex        =   3
         Top             =   390
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frm_lookup_meal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_use_Click()
Dim str_sql As String
Dim pTgl1 As String, pTgl2 As String, str_file As String, kdkaryawan As String
Dim a As New frm_rpt
Dim rs_proses As New ADODB.Recordset
Dim v_proses As Integer
Dim str_param_periode As String

If check_validate_tdbcombo(frm_rpt_summary_salary.TDBCombo_company) = False Then
    MsgBox "No Branch Office selected!", vbInformation, headerMSG
    Exit Sub
End If

'+++++++++++++++++++++++++++++++++ Check Temp Salary Proses ++++++++++++++++++++++++++++++++++++++
str_sql = "SELECT salary_proses FROM temp_sal_proses WHERE company_code = '" & frm_rpt_summary_salary.TDBCombo_company.Text & "'"
rs_proses.Open str_sql, CnG, adOpenForwardOnly, adLockReadOnly
    v_proses = rs_proses!salary_proses
rs_proses.Close
    
    If v_proses = 0 Then
        MsgBox "There are some data that has been changed." & Chr(13) & _
            "Please re salary process!", vbExclamation, headerMSG
        Exit Sub
    End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' str_param1 = Format(DTPicker1.Value, "dd-MMM-yyyy")

If frm_rpt_summary_salary.SSTab1.Tab = 0 Then
    pTgl1 = Format(frm_rpt_summary_salary.DTPicker_monthly.Value, "yyyy-MM-01")
    pTgl2 = Format(frm_rpt_summary_salary.DTPicker_monthly.Value, "yyyy-MM-31")
Else
    pTgl1 = Format(frm_rpt_summary_salary.DTPicker_periode_from.Value, "yyyy-MM-dd")
    pTgl2 = Format(frm_rpt_summary_salary.DTPicker_periode_to.Value, "yyyy-MM-dd")
End If

If Option1(0) Then
    str_file = "\report\rpt_meal.rpt"
Else
    str_file = "\report\rpt_meal_summary.rpt"
End If
  
Dim access As String

access = IIf(LOGIN_LEVEL = 100, "", "AND (a.managerial_access = 0 OR a.managerial_access IS NULL)")

If frm_rpt_summary_salary.cbo_monthly_employee.ListIndex = 0 Then
    str_sql = "SELECT a.employee_code, a.employee_name, a.title_name, a.start_working, " & _
                "'" & Format(frm_rpt_summary_salary.DTPicker_monthly.Value, "yyyy-MM-dd") & "', b.salary_basic, " & _
                "CASE WHEN a.marital_status = 0 THEN 'S' ELSE CONCAT('M/',CAST(IFNULL(a.number_of_children,0) AS CHAR)) END marital_status, " & _
                "fn_jmlHariKerja(a.employee_code,'" & pTgl1 & "','" & pTgl2 & "') jmlHariMasuk," & _
                "IFNULL((SELECT uangmakan FROM m_salary WHERE employee_code = a.employee_code ORDER BY salary_date DESC LIMIT 1),0) uang_makan, " & _
                "a.company_code,'" & LOGIN_FULLNAME & "' prepared, '" & txt_check.Text & "' checked " & _
            "FROM m_employee a LEFT JOIN h_salary_new b ON a.employee_code = b.employee_code " & _
            "WHERE a.company_code = '" & frm_rpt_summary_salary.TDBCombo_company.Text & "' AND " & _
            "a.flag_active <> 0 " & access & " AND " & _
            "RIGHT(b.month,2) = MONTH('" & pTgl1 & "') AND " & _
            "LEFT(b.month,4) = YEAR('" & pTgl1 & "') " & _
            "GROUP BY a.employee_code"
Else
    str_sql = "SELECT a.employee_code, a.employee_name, a.title_name, a.start_working, " & _
                "'" & Format(frm_rpt_summary_salary.DTPicker_monthly.Value, "yyyy-MM-dd") & "', b.salary_basic, " & _
                "CASE WHEN a.marital_status = 0 THEN 'S' ELSE CONCAT('M/',CAST(IFNULL(a.number_of_children,0) AS CHAR)) END marital_status, " & _
                "fn_jmlHariKerja(a.employee_code,'" & pTgl1 & "','" & pTgl2 & "') jmlHariMasuk, " & _
                "IFNULL((SELECT uangmakan FROM m_salary WHERE employee_code = '" & frm_rpt_summary_salary.txt_monthly_employee_code.Text & "' ORDER BY salary_date DESC LIMIT 1),0) uang_makan," & _
                "a.company_code,'" & LOGIN_FULLNAME & "' prepared, '" & txt_check.Text & "' checked " & _
            "FROM m_employee a LEFT JOIN h_salary_new b ON a.employee_code = b.employee_code " & _
            "WHERE a.employee_code = '" & frm_rpt_summary_salary.txt_monthly_employee_code.Text & "' AND " & _
            "a.company_code = '" & frm_rpt_summary_salary.TDBCombo_company.Text & "' AND " & _
            "a.flag_active <> 0 " & access & " AND " & _
            "RIGHT(b.month,2) = MONTH('" & pTgl1 & "') AND " & _
            "LEFT(b.month,4) = YEAR('" & pTgl1 & "') " & _
            "GROUP BY a.employee_code"
End If
        
Call a.Show

If Option1(0) Then
    a.Caption = "SLIP MEAL ALLOWANCE"
Else
    a.Caption = "SUMMARY MEAL ALLOWANCE"
End If

str_param_periode = "MONTH OF " & UCase(Format(pTgl1, "mmmm yyyy"))

Unload Me
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub Form_Load()
Option1(0).Value = True
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
    frm_summary.Visible = False
    txt_check.Text = ""
Else
    frm_summary.Visible = True
End If
End Sub
Private Sub txt_check_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

