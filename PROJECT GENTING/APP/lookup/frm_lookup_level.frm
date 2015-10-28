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
Dim v_Expatriate, v_Expatriate_Calc As String
Dim v_Manager, v_Manager_Calc As String
Dim v_Executive, v_Executive_Calc As String
Dim v_staff, v_Staff_Calc As String
Dim v_NonStaff, v_NonStaff_Calc As String
Dim v_level As String
Public public_int As Integer

Dim rs_employee As New ADODB.Recordset
Private WithEvents poSendMail As clsSendMail
Attribute poSendMail.VB_VarHelpID = -1

Dim pub_employee_code, pub_employee_name, pub_company_code, pub_company_name, pub_email As String
Dim int_sent_false, int_sent_true As Integer
Dim str_smtp, str_sender_mail, str_sender_name As String
Dim str_smtp_port, str_username, str_sender_pwd As String

Private Sub chkExpatriate_Click()
If chkExpatriate.Value = 1 Then
    v_Expatriate = "1"
    v_Expatriate_Calc = "Expatriate"
Else
    v_Expatriate = ""
    v_Expatriate_Calc = ""
End If
End Sub

Private Sub chkManager_Click()
If chkManager.Value = 1 Then
    v_Manager = "2"
    v_Manager_Calc = "Manager"
Else
    v_Manager = ""
    v_Manager_Calc = ""
End If
End Sub

Private Sub chkExecutive_Click()
If chkExecutive.Value = 1 Then
    v_Executive = "3"
    v_Executive_Calc = "Executive"
Else
    v_Executive = ""
    v_Executive_Calc = ""
End If
End Sub

Private Sub chkStaff_Click()
If chkStaff.Value = 1 Then
    v_staff = "4"
    v_Staff_Calc = "Staff"
Else
    v_staff = ""
    v_Staff_Calc = ""
End If
End Sub

Private Sub chkNonStaff_Click()
If chkNonStaff.Value = 1 Then
    v_NonStaff = "5"
    v_NonStaff_Calc = "Non Staff"
Else
    v_NonStaff = ""
    v_NonStaff_Calc = ""
End If
End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub print_recapitulation()
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer
Dim strsql As String
Dim rsemployee As New ADODB.Recordset

int_process = vbNo

str_file = "\report\rpt_recapitulation_salary.rpt" 'rpt_hardys_salary2.rpt"

d1 = Format(frm_rpt_summary_salary.DTPicker_periode_from.Value, "yyyy-MM-dd")
d2 = Format(frm_rpt_summary_salary.DTPicker_periode_to.Value, "yyyy-MM-dd")

str_sql = "call spr_salary_recapitulation('" & d1 & "','" & d2 & "'," _
            & "'" & frm_rpt_summary_salary.TDBCombo_company.Text & "','" & LOGIN_CODE & "'," _
            & "'" & v_Expatriate & "','" & v_Manager & "'," _
            & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "'," _
            & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title & "'," _
            & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_prepare1.Text & "','" & frm_rpt_summary_salary.txt_check1.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_auth11.Text & "','" & frm_rpt_summary_salary.txt_auth21.Text & "')"

Text1 = str_sql

Call a.Show

a.Caption = "SALARY RECAPITULATION"
Call a.rpt_view(str_sql, str_file, str_param_periode)

Unload Me
End Sub

Private Sub print_calculation()
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer
Dim strsql As String
Dim rsemployee As New ADODB.Recordset

int_process = vbNo

str_file = "\report\rpt_salary_calc.rpt" 'rpt_hardys_salary2.rpt"

d1 = Format(frm_rpt_summary_salary.DTPicker2.Value, "yyyy-MM-dd")
d2 = Format(frm_rpt_summary_salary.DTPicker3.Value, "yyyy-MM-dd")

v_level = "Capture from level " & v_Expatriate_Calc & IIf(v_Expatriate_Calc = "", v_Manager_Calc, IIf(v_Manager_Calc <> "", ", " & v_Manager_Calc, v_Manager_Calc)) & _
        IIf(v_Expatriate_Calc = "" And v_Manager_Calc = "", v_Executive_Calc, IIf(v_Executive_Calc <> "", ", " & v_Executive_Calc, v_Executive_Calc)) & _
        IIf(v_Expatriate_Calc = "" And v_Manager_Calc = "" And v_Executive_Calc = "", v_Staff_Calc, IIf(v_Staff_Calc <> "", ", " & v_Staff_Calc, v_Staff_Calc)) & _
        IIf(v_Expatriate_Calc = "" And v_Manager_Calc = "" And v_Executive_Calc = "" And v_Staff_Calc = "", v_NonStaff_Calc, IIf(v_NonStaff_Calc <> "", ", " & v_NonStaff_Calc, v_NonStaff_Calc))
        
str_sql = "call spr_salary_calculation('" & d1 & "','" & d2 & "'," _
            & "'" & frm_rpt_summary_salary.TDBCombo_company.Text & "','" & frm_rpt_summary_salary.txt_address.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_hr_executive.Text & "','" & frm_rpt_summary_salary.txt_title_hr_executive.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_hr_manager.Text & "','" & frm_rpt_summary_salary.txt_title_hr_manager & "'," _
            & "'" & frm_rpt_summary_salary.txt_finance.Text & "','" & frm_rpt_summary_salary.txt_title_finance.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_presdir.Text & "','" & frm_rpt_summary_salary.txt_title_presdir.Text & "'," _
            & "'" & LOGIN_CODE & "','" & v_Expatriate & "','" & v_Manager & "'," _
            & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "','" & v_level & "', " _
            & "'" & frm_rpt_summary_salary.txt_prepare.Text & "','" & frm_rpt_summary_salary.txt_check.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_auth1.Text & "','" & frm_rpt_summary_salary.txt_auth2.Text & "')"
            
str_param_periode = Format(frm_rpt_summary_salary.DTPicker_periode_to.Value, "mmmm yyyy") & ""

Text1 = str_sql

Call a.Show

a.Caption = "PAYMENT CALCULATION"
Call a.rpt_view(str_sql, str_file, str_param_periode)

End Sub

Private Sub print_rfp()
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer
Dim strsql As String
Dim rsemployee As New ADODB.Recordset

int_process = vbNo

str_file = "\report\rpt_rfp.rpt" 'rpt_hardys_salary2.rpt"

d1 = Format(frm_rpt_summary_salary.DTPicker_month_rfp_from, "yyyy-MM-dd")
d2 = Format(frm_rpt_summary_salary.DTPicker_month_rfp_to, "yyyy-MM-dd")

str_sql = "call spr_rfp('" & frm_rpt_summary_salary.TDBCombo_company.Text & "','" & frm_rpt_summary_salary.txt_company_name.Text & "'," _
            & "'" & EMPLOYEE_NAME & "','" & DEPARTMENT_CODE & "'," _
            & "'" & frm_rpt_summary_salary.txt_payment_to.Text & "','" & frm_rpt_summary_salary.txt_payment_for1.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_payment_for2.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_hr_executive2.Text & "','" & frm_rpt_summary_salary.txt_hr_executive2_title.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_hr_manager2.Text & "','" & frm_rpt_summary_salary.txt_hr_manager2_title & "'," _
            & "'" & frm_rpt_summary_salary.txt_finance2.Text & "','" & frm_rpt_summary_salary.txt_finance2_title.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_presdir2.Text & "','" & frm_rpt_summary_salary.txt_presdir2_title.Text & "'," _
            & "'" & v_Expatriate & "','" & v_Manager & "'," _
            & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "','" & d1 & "','" & d2 & "','" & LOGIN_CODE & "', " _
            & "'" & frm_rpt_summary_salary.txt_prepare2.Text & "','" & frm_rpt_summary_salary.txt_check2.Text & "'," _
            & "'" & frm_rpt_summary_salary.txt_auth12.Text & "','" & frm_rpt_summary_salary.txt_auth22.Text & "')"
            
'str_param_periode = Format(DTPicker_periode_to.Value, "mmmm yyyy") & ""

Text1 = str_sql

Call a.Show

a.Caption = "PAYMENT CALCULATION"
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub print_slip(ByVal j As Integer)
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer
Dim strsql As String
Dim rsemployee As New ADODB.Recordset
    
int_process = vbNo

If j = 0 Then
    str_file = "\report\rpt_salary_full.rpt" 'rpt_hardys_salary1.rpt"
ElseIf j = 2 Then
    str_file = "\report\rpt_salary_full.rpt" 'rpt_hardys_salary1.rpt"
ElseIf j = 1 Then
    str_file = "\report\rpt_slip_new.rpt" 'rpt_hardys_salary2.rpt"
End If
d1 = Format(frm_rpt_summary_salary.DTPicker_periode_from.Value, "yyyy-MM-dd")
d2 = Format(frm_rpt_summary_salary.DTPicker_periode_to.Value, "yyyy-MM-dd")

If frm_rpt_summary_salary.cbo_periode_company.ListIndex = 0 Then
    int_flag_company = 0
    str_company_code = "-"
ElseIf frm_rpt_summary_salary.cbo_periode_company.ListIndex = 1 Then
    int_flag_company = 1
    str_company_code = frm_rpt_summary_salary.TDBCombo_company.Columns("company_code").Value
End If

If int_flag_company = 0 Then
    int_flag_employee = 0
    str_employee_code = "-"
ElseIf int_flag_company = 1 Then
    If frm_rpt_summary_salary.cbo_periode_employee.ListIndex = 0 Then
        int_flag_employee = 0
        str_employee_code = "-"
    ElseIf frm_rpt_summary_salary.cbo_periode_employee.ListIndex = 1 Then
        int_flag_employee = 1
        str_employee_code = frm_rpt_summary_salary.txt_periode_employee_code
    End If
End If


    If j = 0 Or j = 2 Then
    '--
        If frm_rpt_summary_salary.cbo_periode_to.ListIndex = 0 Then
            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary.TDBCombo_department.Text & "'," _
                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "'," _
                & "'" & v_Expatriate & "','" & v_Manager & "'," _
                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "', " _
                & "'" & frm_rpt_summary_salary.txt_prepare1.Text & "','" & frm_rpt_summary_salary.txt_check1.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_auth11.Text & "','" & frm_rpt_summary_salary.txt_auth21.Text & "')"
                
            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
        ElseIf frm_rpt_summary_salary.cbo_periode_to.ListIndex = 1 Then
            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary.TDBCombo_department.Text & "'," _
                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "'," _
                & "'" & v_Expatriate & "','" & v_Manager & "'," _
                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "'," _
                & "'" & frm_rpt_summary_salary.txt_prepare1.Text & "','" & frm_rpt_summary_salary.txt_check1.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_auth11.Text & "','" & frm_rpt_summary_salary.txt_auth21.Text & "')"
                
            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
                & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
        End If
    '--
    ElseIf j = 1 Then
    '--
        If frm_rpt_summary_salary.cbo_periode_to.ListIndex = 0 Then
            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary.TDBCombo_department.Text & "'," _
                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "'," _
                & "'" & v_Expatriate & "','" & v_Manager & "'," _
                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "'," _
                & "'" & frm_rpt_summary_salary.txt_prepare1.Text & "','" & frm_rpt_summary_salary.txt_check1.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_auth11.Text & "','" & frm_rpt_summary_salary.txt_auth21.Text & "')"
            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
        ElseIf frm_rpt_summary_salary.cbo_periode_to.ListIndex = 1 Then
            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary.TDBCombo_department.Text & "'," _
                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "'," _
                & "'" & v_Expatriate & "','" & v_Manager & "'," _
                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "'," _
                & "'" & frm_rpt_summary_salary.txt_prepare1.Text & "','" & frm_rpt_summary_salary.txt_check1.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_auth11.Text & "','" & frm_rpt_summary_salary.txt_auth21.Text & "')"
                
            'str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
                & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
        End If
    '--
    End If
    
str_param_periode = Format(frm_rpt_summary_salary.DTPicker1.Value, "mmmm yyyy")


Text1 = str_sql

Call a.Show
a.Caption = "SALARY REPORT"
Call a.rpt_view(str_sql, str_file, str_param_periode)

frm_rpt_summary_salary.fra_process_periode.Visible = False
frm_rpt_summary_salary.fra_periode.Visible = True
End Sub

Private Sub print_slip_ketapang()
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2, dx As Date
Dim int_process As Integer
Dim strsql As String
Dim rsemployee As New ADODB.Recordset
Dim rs_proses As New ADODB.Recordset
Dim v_proses As Integer

'+++++++++++++++++++++++++++++++++ Check Temp Salary Proses ++++++++++++++++++++++++++++++++++++++
str_sql = "SELECT salary_proses FROM temp_sal_proses WHERE company_code = '" & frm_rpt_summary_salary_ketapang.TDBCombo_company.Text & "'"
rs_proses.Open str_sql, CnG, adOpenForwardOnly, adLockReadOnly
    v_proses = rs_proses!salary_proses
rs_proses.Close

    If v_proses = 0 Then
        MsgBox "There are some data that has been changed." & Chr(13) & _
            "Please re salary process!", vbExclamation, headerMSG
        Exit Sub
    End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

int_process = vbNo

str_file = "\report\rpt_slip_new.rpt" 'rpt_hardys_salary2.rpt"

d1 = Format(frm_rpt_summary_salary_ketapang.DTPicker_periode_from.Value, "yyyy-MM-dd")
d2 = Format(frm_rpt_summary_salary_ketapang.DTPicker_periode_to.Value, "yyyy-MM-dd")

If frm_rpt_summary_salary_ketapang.cbo_periode_company.ListIndex = 0 Then
    int_flag_company = 0
    str_company_code = "-"
ElseIf frm_rpt_summary_salary_ketapang.cbo_periode_company.ListIndex = 1 Then
    int_flag_company = 1
    str_company_code = frm_rpt_summary_salary_ketapang.TDBCombo_company.Columns("company_code").Value
End If

If int_flag_company = 0 Then
    int_flag_employee = 0
    str_employee_code = "-"
ElseIf int_flag_company = 1 Then
    If frm_rpt_summary_salary_ketapang.cbo_periode_employee.ListIndex = 0 Then
        int_flag_employee = 0
        str_employee_code = "-"
    ElseIf frm_rpt_summary_salary_ketapang.cbo_periode_employee.ListIndex = 1 Then
        int_flag_employee = 1
        str_employee_code = frm_rpt_summary_salary_ketapang.txt_periode_employee_code
    End If
End If

'--
    If frm_rpt_summary_salary_ketapang.cbo_periode_to.ListIndex = 0 Then
            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary_ketapang.TDBCombo_department.Text & "'," _
                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary_ketapang.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
                & "'" & frm_rpt_summary_salary_ketapang.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_hr_executive1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary_ketapang.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_hr_manager1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary_ketapang.txt_finance1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_finance1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary_ketapang.txt_presdir1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_presdir1_title.Text & "'," _
                & "'" & v_Expatriate & "','" & v_Manager & "'," _
                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "')"
        ElseIf frm_rpt_summary_salary.cbo_periode_to.ListIndex = 1 Then
            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                & int_flag_company & ",'" & str_company_code & "','" & frm_rpt_summary_salary_ketapang.TDBCombo_department.Text & "'," _
                & int_flag_employee & ",'" & str_employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary_ketapang.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
                & "'" & frm_rpt_summary_salary_ketapang.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_hr_executive1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary_ketapang.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_hr_manager1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary_ketapang.txt_finance1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_finance1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary_ketapang.txt_presdir1.Text & "','" & frm_rpt_summary_salary_ketapang.txt_presdir1_title.Text & "'," _
                & "'" & v_Expatriate & "','" & v_Manager & "'," _
                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "')"
    End If
'--


Text1 = str_sql

Call a.Show

a.Caption = "SALARY SLIP REPORT"
Call a.rpt_view(str_sql, str_file, str_param_periode)

frm_rpt_summary_salary_ketapang.fra_process_periode.Visible = False
frm_rpt_summary_salary_ketapang.fra_periode.Visible = True
End Sub

Private Sub print_slip_email()
Dim str_sql, str_sql2, str_sql3, str_param_periode, str_file, str1, str_file_out As String
Dim int_flag_department As Integer, str_department_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim dt1, dt2 As Date
Dim d1, d2 As String
Dim a As New frm_rpt
Dim i As Integer
Dim approve, correct As String
Dim rs_proses As New ADODB.Recordset
Dim v_proses As Integer
    
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
    i = MsgBox("Send Salary Slip Employee mail ? ", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    int_sent_false = 0
    int_sent_true = 0
    
    str_file = "\report\rpt_slip_new.rpt"
    
    Dim tgl1 As String, tgl2 As String, kdkaryawan As String
    
    tgl1 = Format(frm_rpt_summary_salary.DTPicker_periode_from.Value, "yyyy-MM-dd")
    tgl2 = Format(frm_rpt_summary_salary.DTPicker_periode_to.Value, "yyyy-MM-dd")
    kdkaryawan = frm_rpt_summary_salary.txt_periode_employee_code.Text
    
    If frm_rpt_summary_salary.cbo_periode_department.ListIndex = 0 Then
        int_flag_department = 0
        str_department_code = "-"
    ElseIf frm_rpt_summary_salary.cbo_periode_department.ListIndex = 1 Then
        int_flag_department = 1
        str_department_code = TDBCombo_department.Text
    End If
    
    If int_flag_department = 0 Then
        int_flag_employee = 0
        str_employee_code = "-"
        str1 = "SELECT a.* " _
            & "FROM m_employee a JOIN h_attendance b ON a.employee_code = b.employee_code " _
            & "WHERE (a.email <> CASE WHEN a.email = '-' THEN '-' ELSE '' END) AND ifnull(a.flag_active,0) = 1 AND a.company_code = '" & frm_rpt_summary_salary.TDBCombo_company.Text & "' " _
            & "Group By a.employee_code"
    ElseIf int_flag_department = 1 Then
        If frm_rpt_summary_salary.cbo_periode_employee.ListIndex = 0 Then
            int_flag_employee = 0
            str_employee_code = "-"
        str1 = "SELECT a.* " _
            & "FROM m_employee a JOIN h_attendance b ON a.employee_code = b.employee_code " _
            & "WHERE (a.email <> CASE WHEN a.email = '-' THEN '-' ELSE '' END) AND ifnull(a.flag_active,0) = 1 AND a.company_code = '" & frm_rpt_summary_salary.TDBCombo_company.Text & "' " _
            & "AND a.department_code = '" & str_department_code & "' Group By a.employee_code"
    ElseIf cbo_periode_employee.ListIndex = 1 Then
            int_flag_employee = 1
            str_employee_code = frm_rpt_summary_salary.txt_periode_employee_code
            str1 = "SELECT a.* " _
                & "FROM m_employee a JOIN h_attendance b ON a.employee_code = b.employee_code " _
                & "WHERE (a.email <> CASE WHEN a.email = '-' THEN '-' ELSE '' END) AND ifnull(a.flag_active,0) = 1 AND a.company_code = '" & frm_rpt_summary_salary.TDBCombo_company.Text & "' " _
                & "AND a.employee_code = '" & str_employee_code & "' Group By a.employee_code"
        End If
    End If
    
    If rs_employee.State = 1 Then rs_employee.Close
    rs_employee.Open str1, CnG, adOpenStatic, adLockReadOnly
    If rs_employee.RecordCount > 0 Then rs_employee.MoveFirst
    
    d1 = Format(frm_rpt_summary_salary.DTPicker_periode_from.Value, "yyyy-MM-dd")
    d2 = Format(frm_rpt_summary_salary.DTPicker_periode_to.Value, "yyyy-MM-dd")
    
    If rs_employee.RecordCount <= 0 Then
        rs_employee.Close
        MsgBox "This month data has not been processed!" & Chr(13) & "Please Access Utility - Salary process First"
        Exit Sub
    End If
    
    If rs_employee.RecordCount > 0 Then
        frm_rpt_summary_salary.ProgressBar2.Visible = True
        frm_rpt_summary_salary.Label23.Visible = True
        frm_rpt_summary_salary.Label22.Visible = True
        frm_rpt_summary_salary.ProgressBar2.Value = 0
        frm_rpt_summary_salary.ProgressBar2.Max = rs_employee.RecordCount
        frm_rpt_summary_salary.fra_periode.Enabled = False
        
        frm_rpt_summary_salary.cmd_print_slip.Enabled = False
        frm_rpt_summary_salary.cmdPrint.Enabled = False
        frm_rpt_summary_salary.cmdPrint_Rekap.Enabled = False
        frm_rpt_summary_salary.cmd_send_mail.Enabled = False
    End If
    
    keluar = False
    While Not rs_employee.EOF
        If keluar = True Then
            rs_employee.Close
            Exit Sub
        End If
        
        frm_rpt_summary_salary.ProgressBar2.Value = frm_rpt_summary_salary.ProgressBar2.Value + 1
        frm_rpt_summary_salary.ProgressBar2.Value = frm_rpt_summary_salary.ProgressBar2.Value
        frm_rpt_summary_salary.Label22.Caption = "Proses " & frm_rpt_summary_salary.ProgressBar2.Value & "/" & rs_employee.RecordCount
        
        If frm_rpt_summary_salary.cbo_periode_to.ListIndex = 0 Then
            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                & "0,'" & frm_rpt_summary_salary.TDBCombo_company & "','" & frm_rpt_summary_salary.TDBCombo_department.Text & "'," _
                & "1,'" & rs_employee!employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "'," _
                & "'" & v_Expatriate & "','" & v_Manager & "'," _
                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "'," _
                & "'" & frm_rpt_summary_salary.txt_prepare1.Text & "','" & frm_rpt_summary_salary.txt_check1.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_auth11.Text & "','" & frm_rpt_summary_salary.txt_auth21.Text & "')"
        ElseIf frm_rpt_summary_salary.cbo_periode_to.ListIndex = 1 Then
            str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," _
                & "0,'" & frm_rpt_summary_salary.TDBCombo_company & "','" & frm_rpt_summary_salary.TDBCombo_department.Text & "'," _
                & "1,'" & rs_employee!employee_code & "','" & DATA_LEVEL & "'," & frm_rpt_summary_salary.cbo_periode_department.ListIndex & ",'" & LOGIN_CODE & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_executive1.Text & "','" & frm_rpt_summary_salary.txt_hr_executive1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_hr_manager1.Text & "','" & frm_rpt_summary_salary.txt_hr_manager1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_finance1.Text & "','" & frm_rpt_summary_salary.txt_finance1_title.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_presdir1.Text & "','" & frm_rpt_summary_salary.txt_presdir1_title.Text & "'," _
                & "'" & v_Expatriate & "','" & v_Manager & "'," _
                & "'" & v_Executive & "','" & v_staff & "','" & v_NonStaff & "'," _
                & "'" & frm_rpt_summary_salary.txt_prepare1.Text & "','" & frm_rpt_summary_salary.txt_check1.Text & "'," _
                & "'" & frm_rpt_summary_salary.txt_auth11.Text & "','" & frm_rpt_summary_salary.txt_auth21.Text & "')"
        End If
        
        str_param_periode = "PERIODE : (" & Format(frm_rpt_summary_salary.DTPicker1.Value, "yyyy-MM") & ")"
    
    
        '-- creating pdf
        str_file_out = App.Path & "\mail\slip_" & Format(frm_rpt_summary_salary.DTPicker1, "yyyymm") & "_" & rs_employee!employee_code & ".pdf"
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
    
    rs_employee.Close
    Call show_msg
    
    frm_rpt_summary_salary.ProgressBar2.Visible = False
    frm_rpt_summary_salary.Label22.Visible = False
    frm_rpt_summary_salary.Label23.Visible = False
    frm_rpt_summary_salary.fra_periode.Enabled = True
    
    frm_rpt_summary_salary.cmd_print_slip.Enabled = True
    frm_rpt_summary_salary.cmdPrint_Rekap.Enabled = True
    frm_rpt_summary_salary.cmdPrint.Enabled = True
    frm_rpt_summary_salary.cmd_send_mail.Enabled = True
End Sub

Private Sub cmd_use_Click()

If public_int = 1 Then
    Call print_recapitulation
ElseIf public_int = 2 Then
    Call print_calculation
ElseIf public_int = 3 Then
    Call print_rfp
ElseIf public_int = 4 Then
    Call print_slip(1)
ElseIf public_int = 5 Then
    Call print_slip(0)
ElseIf public_int = 6 Then
    Call print_slip_ketapang
ElseIf public_int = 7 Then
    Unload Me
    Call print_slip_email
End If

Unload Me
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim strsql As String
Dim v_allow_access As Integer

Call load_data_setting_mail

strsql = "SELECT * from t_user_access_level " _
    & "WHERE level_code = '" & LOGIN_CODE & "' AND access_level_code = 1"
rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
    v_allow_access = rs!allow_access
    If v_allow_access = 0 Then
        chkExpatriate.Enabled = False
    Else
        chkExpatriate.Enabled = True
    End If
    chkExpatriate.Value = IIf(v_allow_access = -1, 1, v_allow_access)
End If
rs.Close

strsql = "SELECT * from t_user_access_level " _
    & "WHERE level_code = '" & LOGIN_CODE & "' AND access_level_code = 2"
rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
    v_allow_access = rs!allow_access
    If v_allow_access = 0 Then
        chkManager.Enabled = False
    Else
        chkManager.Enabled = True
    End If
    chkManager.Value = IIf(v_allow_access = -1, 1, v_allow_access)
End If
rs.Close

strsql = "SELECT * from t_user_access_level " _
    & "WHERE level_code = '" & LOGIN_CODE & "' AND access_level_code = 3"
rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
    v_allow_access = rs!allow_access
    If v_allow_access = 0 Then
        chkExecutive.Enabled = False
    Else
        chkExecutive.Enabled = True
    End If
    chkExecutive.Value = IIf(v_allow_access, 1, v_allow_access)
End If
rs.Close

strsql = "SELECT * from t_user_access_level " _
    & "WHERE level_code = '" & LOGIN_CODE & "' AND access_level_code = 4"
rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
    v_allow_access = rs!allow_access
    If v_allow_access = 0 Then
        chkStaff.Enabled = False
    Else
        chkStaff.Enabled = True
    End If
    chkStaff.Value = IIf(v_allow_access, 1, v_allow_access)
End If
rs.Close

strsql = "SELECT * from t_user_access_level " _
    & "WHERE level_code = '" & LOGIN_CODE & "' AND access_level_code = 5"
rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
    v_allow_access = rs!allow_access
    If v_allow_access = 0 Then
        chkNonStaff.Enabled = False
    Else
        chkNonStaff.Enabled = True
    End If
    chkNonStaff.Value = IIf(v_allow_access, 1, v_allow_access)
End If
rs.Close
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
        .SMTPPort = str_smtp_port
        '.SMTPHost = "mail.solusisentraldata.com"
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
    '        .Username = "send.ktc@solusisentraldata.com"  ' Optional, default = Null String
    '        .Password = "sepatukuda2011"                            ' Optional, default = Null String, value is NOT saved
            .Username = str_username                    ' Optional, default = Null String
            .Password = str_sender_pwd
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

Private Sub show_msg()
    MsgBox "There are " & int_sent_true & " mail are sent successfully!" & vbCrLf _
        & int_sent_false & " are fails!" & vbCrLf _
        & "For more detail info, let see the 'log mail'!", vbInformation, headerMSG
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
    .Fields("description").Value = "sent failed"
'    .Fields("description").Value = "Your attempt to send mail failed for the following reason(s): " _
'                                    & vbCrLf & Explanation
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

Private Sub load_data_setting_mail()
Dim rs1 As New ADODB.Recordset
Dim strsql As String

strsql = "select * from s_mail where s_number = 1"
rs1.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs1.RecordCount > 0 Then
    str_sender_mail = rs1!s_sender_email
    str_sender_name = rs1!s_sender_name
    str_smtp = rs1!smtp
    str_smtp_port = rs1!PORT
    str_username = rs1!Username
    str_sender_pwd = RC4DeCryptASC(rs1!Password, pEncryptionPassword)
Else
    str_sender_mail = ""
    str_sender_name = ""
    str_smtp = ""
    str_smtp_port = ""
    str_username = ""
    str_sender_pwd = ""
End If
End Sub
