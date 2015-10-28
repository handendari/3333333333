VERSION 5.00
Begin VB.Form frm_lookup_level 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SELECT LEVEL"
   ClientHeight    =   3510
   ClientLeft      =   -15
   ClientTop       =   345
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_close 
      Caption         =   "&Close"
      Height          =   645
      Left            =   3450
      Picture         =   "frm_lookup_level.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmd_use 
      Caption         =   "&OK"
      Height          =   645
      Left            =   2400
      Picture         =   "frm_lookup_level.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECT LEVEL TO REPORT"
      Height          =   2475
      Left            =   360
      TabIndex        =   0
      Top             =   150
      Width           =   4065
      Begin VB.CheckBox chkNonStaff 
         Caption         =   "Non Staff"
         Height          =   225
         Left            =   270
         TabIndex        =   5
         Top             =   1950
         Width           =   1455
      End
      Begin VB.CheckBox chkStaff 
         Caption         =   "Staff"
         Height          =   225
         Left            =   270
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox chkExecutive 
         Caption         =   "Executive"
         Height          =   225
         Left            =   270
         TabIndex        =   3
         Top             =   1170
         Width           =   1455
      End
      Begin VB.CheckBox chkExpatriate 
         Caption         =   "Expatriate"
         Height          =   225
         Left            =   270
         TabIndex        =   2
         Top             =   420
         Width           =   1455
      End
      Begin VB.CheckBox chkManager 
         Caption         =   "Manager"
         Height          =   225
         Left            =   270
         TabIndex        =   1
         Top             =   780
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_lookup_level"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim v_Expatriate, v_Expatriate_Calc As String
'Dim v_Manager, v_Manager_Calc As String
'Dim v_Executive, v_Executive_Calc As String
'Dim v_staff, v_Staff_Calc As String
'Dim v_NonStaff, v_NonStaff_Calc As String
'Dim v_level As String
'Public public_int As Integer
'
'Private Sub chkExpatriate_Click()
'If chkExpatriate.Value = 1 Then
'    v_Expatriate = "1"
'    v_Expatriate_Calc = "Expatriate"
'Else
'    v_Expatriate = ""
'    v_Expatriate_Calc = ""
'End If
'End Sub
'
'Private Sub chkManager_Click()
'If chkManager.Value = 1 Then
'    v_Manager = "2"
'    v_Manager_Calc = "Manager"
'Else
'    v_Manager = ""
'    v_Manager_Calc = ""
'End If
'End Sub
'
'Private Sub chkExecutive_Click()
'If chkExecutive.Value = 1 Then
'    v_Executive = "3"
'    v_Executive_Calc = "Executive"
'Else
'    v_Executive = ""
'    v_Executive_Calc = ""
'End If
'End Sub
'
'Private Sub chkStaff_Click()
'If chkStaff.Value = 1 Then
'    v_staff = "4"
'    v_Staff_Calc = "Staff"
'Else
'    v_staff = ""
'    v_Staff_Calc = ""
'End If
'End Sub
'
'Private Sub chkNonStaff_Click()
'If chkNonStaff.Value = 1 Then
'    v_NonStaff = "5"
'    v_NonStaff_Calc = "Non Staff"
'Else
'    v_NonStaff = ""
'    v_NonStaff_Calc = ""
'End If
'End Sub
'
'Private Sub cmd_close_Click()
'Unload Me
'End Sub
'
'Private Sub print_recapitulation()
'Dim str_sql, str_param_periode, str_file, str1, str2  As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim a As New frm_rpt
'Dim d1, d2, dx As Date
'Dim int_process As Integer
'Dim strSQl As String
'Dim rsemployee As New ADODB.Recordset
'
'int_process = vbNo
'
'str_file = "\report\rpt_recapitulation_salary.rpt" 'rpt_hardys_salary2.rpt"
'
'd1 = Format(frm_rpt_summary_salary.DTPicker_periode_from.Value, "yyyy-MM-dd")
'd2 = Format(frm_rpt_summary_salary.DTPicker_periode_to.Value, "yyyy-MM-dd")
'
'str_sql = "call spr_salary_recapitulation('" & d1 & "','" & d2 & "'," _
'            & "'" & frm_rpt_summary_salary.TDBCombo_company.Text & "','" & LOGIN_CODE & "'," _
'            & "'" & v_Expatriate & "','" & v_Manager & "'," _
'            & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "'," _
'            & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
'            & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title & "'," _
'            & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
'            & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "')"
'
'Text1 = str_sql
'
'Call a.Show
'
'a.Caption = "SALARY RECAPITULATION"
'Call a.rpt_view(str_sql, str_file, str_param_periode)
'
'Unload Me
'End Sub
'
'Private Sub print_calculation()
'Dim str_sql, str_param_periode, str_file, str1, str2  As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim a As New frm_rpt
'Dim d1, d2, dx As Date
'Dim int_process As Integer
'Dim strSQl As String
'Dim rsemployee As New ADODB.Recordset
'
'int_process = vbNo
'
'str_file = "\report\rpt_salary_calc.rpt" 'rpt_hardys_salary2.rpt"
'
'd1 = Format(frm_rpt_summary_salary.DTPicker2.Value, "yyyy-MM-dd")
'd2 = Format(frm_rpt_summary_salary.DTPicker3.Value, "yyyy-MM-dd")
'
'v_level = "Capture from level " & v_Expatriate_Calc & IIf(v_Expatriate_Calc = "", v_Manager_Calc, IIf(v_Manager_Calc <> "", ", " & v_Manager_Calc, v_Manager_Calc)) & _
'        IIf(v_Expatriate_Calc = "" And v_Manager_Calc = "", v_Executive_Calc, IIf(v_Executive_Calc <> "", ", " & v_Executive_Calc, v_Executive_Calc)) & _
'        IIf(v_Expatriate_Calc = "" And v_Manager_Calc = "" And v_Executive_Calc = "", v_Staff_Calc, IIf(v_Staff_Calc <> "", ", " & v_Staff_Calc, v_Staff_Calc)) & _
'        IIf(v_Expatriate_Calc = "" And v_Manager_Calc = "" And v_Executive_Calc = "" And v_Staff_Calc = "", v_NonStaff_Calc, IIf(v_NonStaff_Calc <> "", ", " & v_NonStaff_Calc, v_NonStaff_Calc))
'
'str_sql = "call spr_salary_calculation('" & d1 & "','" & d2 & "'," _
'            & "'" & frm_rpt_summary_salary.TDBCombo_company.Text & "','" & frm_rpt_summary_salary.txt_address.Text & "'," _
'            & "'" & frm_rpt_summary_salary.txt_hr_executive.Text & "','" & frm_rpt_summary_salary.txt_title_hr_executive.Text & "'," _
'            & "'" & frm_rpt_summary_salary.txt_hr_manager.Text & "','" & frm_rpt_summary_salary.txt_title_hr_manager & "'," _
'            & "'" & frm_rpt_summary_salary.txt_finance.Text & "','" & frm_rpt_summary_salary.txt_title_finance.Text & "'," _
'            & "'" & frm_rpt_summary_salary.txt_presdir.Text & "','" & frm_rpt_summary_salary.txt_title_presdir.Text & "'," _
'            & "'" & LOGIN_CODE & "','" & v_Expatriate & "','" & v_Manager & "'," _
'            & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "','" & v_level & "')"
'
'str_param_periode = Format(frm_rpt_summary_salary.DTPicker_periode_to.Value, "mmmm yyyy") & ""
'
'Text1 = str_sql
'
'Call a.Show
'
'a.Caption = "PAYMENT CALCULATION"
'Call a.rpt_view(str_sql, str_file, str_param_periode)
'
'End Sub
'
'Private Sub print_rfp()
'Dim str_sql, str_param_periode, str_file, str1, str2  As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim a As New frm_rpt
'Dim d1, d2, dx As Date
'Dim int_process As Integer
'Dim strSQl As String
'Dim rsemployee As New ADODB.Recordset
'
'int_process = vbNo
'
'str_file = "\report\rpt_rfp.rpt" 'rpt_hardys_salary2.rpt"
'
'd1 = Format(frm_rpt_summary_salary.DTPicker_month_rfp_from, "yyyy-MM-dd")
'd2 = Format(frm_rpt_summary_salary.DTPicker_month_rfp_to, "yyyy-MM-dd")
'
'str_sql = "call spr_rfp('" & frm_rpt_summary_salary.TDBCombo_company.Text & "','" & frm_rpt_summary_salary.txt_company_name.Text & "'," _
'            & "'" & EMPLOYEE_NAME & "','" & DEPARTMENT_CODE & "'," _
'            & "'" & frm_rpt_summary_salary.txt_payment_to.Text & "','" & frm_rpt_summary_salary.txt_payment_for1.Text & "'," _
'            & "'" & frm_rpt_summary_salary.txt_payment_for2.Text & "'," _
'            & "'" & frm_rpt_summary_salary.txt_hr_executive2.Text & "','" & frm_rpt_summary_salary.txt_hr_executive2_title.Text & "'," _
'            & "'" & frm_rpt_summary_salary.txt_hr_manager2.Text & "','" & frm_rpt_summary_salary.txt_hr_manager2_title & "'," _
'            & "'" & frm_rpt_summary_salary.txt_finance2.Text & "','" & frm_rpt_summary_salary.txt_finance2_title.Text & "'," _
'            & "'" & frm_rpt_summary_salary.txt_presdir2.Text & "','" & frm_rpt_summary_salary.txt_presdir2_title.Text & "'," _
'            & "'" & v_Expatriate & "','" & v_Manager & "'," _
'            & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "','" & d1 & "','" & d2 & "','" & LOGIN_CODE & "')"
'
''str_param_periode = Format(DTPicker_periode_to.Value, "mmmm yyyy") & ""
'
'Text1 = str_sql
'
'Call a.Show
'
'a.Caption = "PAYMENT CALCULATION"
'Call a.rpt_view(str_sql, str_file, str_param_periode)
'End Sub
'
'Private Sub print_slip(ByVal j As Integer)
'Dim str_sql, str_param_periode, str_file, str1, str2  As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim a As New frm_rpt
'Dim d1, d2, dx As Date
'Dim int_process As Integer
'Dim strSQl As String
'Dim rsemployee As New ADODB.Recordset
'
'int_process = vbNo
'
'If j = 0 Then
'    str_file = "\report\rpt_summary_salary.rpt" 'rpt_hardys_salary1.rpt"
'ElseIf j = 2 Then
'    str_file = "\report\rpt_summary_salary.rpt" 'rpt_hardys_salary1.rpt"
'ElseIf j = 1 Then
'    str_file = "\report\rpt_slip_salary.rpt" 'rpt_hardys_salary2.rpt"
'End If
'd1 = Format(frm_rpt_summary_salary.DTPicker_periode_from.Value, "yyyy-MM-dd")
'd2 = Format(frm_rpt_summary_salary.DTPicker_periode_to.Value, "yyyy-MM-dd")
'
'If frm_rpt_summary_salary.cbo_periode_company.ListIndex = 0 Then
'    int_flag_company = 0
'    str_company_code = "-"
'ElseIf frm_rpt_summary_salary.cbo_periode_company.ListIndex = 1 Then
'    int_flag_company = 1
'    str_company_code = frm_rpt_summary_salary.TDBCombo_company.Columns("company_code").Value
'End If
'
'If int_flag_company = 0 Then
'    int_flag_employee = 0
'    str_employee_code = "-"
'ElseIf int_flag_company = 1 Then
'    If frm_rpt_summary_salary.cbo_periode_employee.ListIndex = 0 Then
'        int_flag_employee = 0
'        str_employee_code = "-"
'    ElseIf frm_rpt_summary_salary.cbo_periode_employee.ListIndex = 1 Then
'        int_flag_employee = 1
'        str_employee_code = frm_rpt_summary_salary.txt_periode_employee_code
'    End If
'End If
'
'
'    If j = 0 Or j = 2 Then
'    '--
'        If frm_rpt_summary_salary.cbo_periode_to.ListIndex = 0 Then
'            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
'                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary.TDBCombo_department.Text & "'," _
'                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
'                & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "'," _
'                & "'" & v_Expatriate & "','" & v_Manager & "'," _
'                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "')"
'
'            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
'        ElseIf frm_rpt_summary_salary.cbo_periode_to.ListIndex = 1 Then
'            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
'                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary.TDBCombo_department.Text & "'," _
'                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
'                & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "'," _
'                & "'" & v_Expatriate & "','" & v_Manager & "'," _
'                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "')"
'
'            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
'                & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
'        End If
'    '--
'    ElseIf j = 1 Then
'    '--
'        If frm_rpt_summary_salary.cbo_periode_to.ListIndex = 0 Then
'            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
'                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary.TDBCombo_department.Text & "'," _
'                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
'                & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "'," _
'                & "'" & v_Expatriate & "','" & v_Manager & "'," _
'                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "')"
'            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
'        ElseIf frm_rpt_summary_salary.cbo_periode_to.ListIndex = 1 Then
'            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
'                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary.TDBCombo_department.Text & "'," _
'                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
'                & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "'," _
'                & "'" & v_Expatriate & "','" & v_Manager & "'," _
'                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "')"
'
'            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
'                & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
'        End If
'    '--
'    End If
'
'str_param_periode = Format(frm_rpt_summary_salary.DTPicker1.Value, "mmmm yyyy")
'
'
'Text1 = str_sql
'
'Call a.Show
'a.Caption = "SALARY REPORT"
'Call a.rpt_view(str_sql, str_file, str_param_periode)
'
'frm_rpt_summary_salary.fra_process_periode.Visible = False
'frm_rpt_summary_salary.fra_periode.Visible = True
'End Sub
'
'Private Sub print_slip_ketapang()
'Dim str_sql, str_param_periode, str_file, str1, str2  As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim a As New frm_rpt
'Dim d1, d2, dx As Date
'Dim int_process As Integer
'Dim strSQl As String
'Dim rsemployee As New ADODB.Recordset
'Dim rs_proses As New ADODB.Recordset
'Dim v_proses As Integer
'
''+++++++++++++++++++++++++++++++++ Check Temp Salary Proses ++++++++++++++++++++++++++++++++++++++
'str_sql = "SELECT salary_proses FROM temp_sal_proses WHERE company_code = '" & frm_rpt_summary_salary_ketapang.TDBCombo_company.Text & "'"
'rs_proses.Open str_sql, CnG, adOpenForwardOnly, adLockReadOnly
'    v_proses = rs_proses!salary_proses
'rs_proses.Close
'
'    If v_proses = 0 Then
'        MsgBox "There are some data that has been changed." & Chr(13) & _
'            "Please re salary process!", vbExclamation, headerMSG
'        Exit Sub
'    End If
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'int_process = vbNo
'
'str_file = "\report\rpt_slip_new.rpt" 'rpt_hardys_salary2.rpt"
'
'd1 = Format(frm_rpt_summary_salary_ketapang.DTPicker_periode_from.Value, "yyyy-MM-dd")
'd2 = Format(frm_rpt_summary_salary_ketapang.DTPicker_periode_to.Value, "yyyy-MM-dd")
'
'If frm_rpt_summary_salary_ketapang.cbo_periode_company.ListIndex = 0 Then
'    int_flag_company = 0
'    str_company_code = "-"
'ElseIf frm_rpt_summary_salary_ketapang.cbo_periode_company.ListIndex = 1 Then
'    int_flag_company = 1
'    str_company_code = frm_rpt_summary_salary_ketapang.TDBCombo_company.Columns("company_code").Value
'End If
'
'If int_flag_company = 0 Then
'    int_flag_employee = 0
'    str_employee_code = "-"
'ElseIf int_flag_company = 1 Then
'    If frm_rpt_summary_salary_ketapang.cbo_periode_employee.ListIndex = 0 Then
'        int_flag_employee = 0
'        str_employee_code = "-"
'    ElseIf frm_rpt_summary_salary_ketapang.cbo_periode_employee.ListIndex = 1 Then
'        int_flag_employee = 1
'        str_employee_code = frm_rpt_summary_salary_ketapang.txt_periode_employee_code
'    End If
'End If
'
''--
'    If frm_rpt_summary_salary_ketapang.cbo_periode_to.ListIndex = 0 Then
'            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
'                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary_ketapang.TDBCombo_department.Text & "'," _
'                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary_ketapang.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
'                & "'" & frm_rpt_summary_salary_ketapang.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_hr_executive1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary_ketapang.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_hr_manager1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary_ketapang.txt_finance1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_finance1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary_ketapang.txt_presdir1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_presdir1_title.Text & "'," _
'                & "'" & v_Expatriate & "','" & v_Manager & "'," _
'                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "')"
'        ElseIf frm_rpt_summary_salary.cbo_periode_to.ListIndex = 1 Then
'            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
'                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary_ketapang.TDBCombo_department.Text & "'," _
'                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary_ketapang.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
'                & "'" & frm_rpt_summary_salary_ketapang.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_hr_executive1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary_ketapang.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_hr_manager1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary_ketapang.txt_finance1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_finance1_title.Text & "'," _
'                & "'" & frm_rpt_summary_salary_ketapang.txt_presdir1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_presdir1_title.Text & "'," _
'                & "'" & v_Expatriate & "','" & v_Manager & "'," _
'                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "')"
'    End If
''--
'
'
'Text1 = str_sql
'
'Call a.Show
'
'a.Caption = "SALARY SLIP REPORT"
'Call a.rpt_view(str_sql, str_file, str_param_periode)
'
'frm_rpt_summary_salary_ketapang.fra_process_periode.Visible = False
'frm_rpt_summary_salary_ketapang.fra_periode.Visible = True
'End Sub
'
'Private Sub cmd_use_Click()
'
'If public_int = 1 Then
'    Call print_recapitulation
'ElseIf public_int = 2 Then
'    Call print_calculation
'ElseIf public_int = 3 Then
'    Call print_rfp
'ElseIf public_int = 4 Then
'    Call print_slip(1)
'ElseIf public_int = 5 Then
'    Call print_slip(0)
'ElseIf public_int = 6 Then
'    Call print_slip_ketapang
'End If
'
'End Sub
'
'Private Sub Form_Load()
'Dim rs As New ADODB.Recordset
'Dim strSQl As String
'Dim v_allow_access As Integer
'
'strSQl = "SELECT * from t_user_access_level " _
'    & "WHERE level_code = '" & LOGIN_CODE & "' AND access_level_code = 1"
'rs.Open strSQl, CnG, adOpenForwardOnly, adLockReadOnly
'
'If rs.RecordCount > 0 Then
'    v_allow_access = rs!allow_access
'    If v_allow_access = 0 Then
'        chkExpatriate.Enabled = False
'    Else
'        chkExpatriate.Enabled = True
'    End If
'    chkExpatriate.Value = IIf(v_allow_access = -1, 1, v_allow_access)
'End If
'rs.Close
'
'strSQl = "SELECT * from t_user_access_level " _
'    & "WHERE level_code = '" & LOGIN_CODE & "' AND access_level_code = 2"
'rs.Open strSQl, CnG, adOpenForwardOnly, adLockReadOnly
'
'If rs.RecordCount > 0 Then
'    v_allow_access = rs!allow_access
'    If v_allow_access = 0 Then
'        chkManager.Enabled = False
'    Else
'        chkManager.Enabled = True
'    End If
'    chkManager.Value = IIf(v_allow_access = -1, 1, v_allow_access)
'End If
'rs.Close
'
'strSQl = "SELECT * from t_user_access_level " _
'    & "WHERE level_code = '" & LOGIN_CODE & "' AND access_level_code = 3"
'rs.Open strSQl, CnG, adOpenForwardOnly, adLockReadOnly
'
'If rs.RecordCount > 0 Then
'    v_allow_access = rs!allow_access
'    If v_allow_access = 0 Then
'        chkExecutive.Enabled = False
'    Else
'        chkExecutive.Enabled = True
'    End If
'    chkExecutive.Value = IIf(v_allow_access, 1, v_allow_access)
'End If
'rs.Close
'
'strSQl = "SELECT * from t_user_access_level " _
'    & "WHERE level_code = '" & LOGIN_CODE & "' AND access_level_code = 4"
'rs.Open strSQl, CnG, adOpenForwardOnly, adLockReadOnly
'
'If rs.RecordCount > 0 Then
'    v_allow_access = rs!allow_access
'    If v_allow_access = 0 Then
'        chkStaff.Enabled = False
'    Else
'        chkStaff.Enabled = True
'    End If
'    chkStaff.Value = IIf(v_allow_access, 1, v_allow_access)
'End If
'rs.Close
'
'strSQl = "SELECT * from t_user_access_level " _
'    & "WHERE level_code = '" & LOGIN_CODE & "' AND access_level_code = 5"
'rs.Open strSQl, CnG, adOpenForwardOnly, adLockReadOnly
'
'If rs.RecordCount > 0 Then
'    v_allow_access = rs!allow_access
'    If v_allow_access = 0 Then
'        chkNonStaff.Enabled = False
'    Else
'        chkNonStaff.Enabled = True
'    End If
'    chkNonStaff.Value = IIf(v_allow_access, 1, v_allow_access)
'End If
'rs.Close
'End Sub
'Private Sub cmd_use_Click()
'
'End Sub
