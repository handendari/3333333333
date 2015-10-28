VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rsemployee As New ADODB.Recordset
Dim strSQl As String

Public Sub insert_m_sallary(pDepartment As String, pTgl1 As String, pTgl2 As String)
'INSERT INTO h_salary
'    (
'    MONTH,
'    employee_code,
'    salary_code,
'    company_code,
'    salary_name,
'    date_from,
'    date_to,
'    flag_main_salary,
'    flag_sign,
'    flag_detail,
'    flag_use_formula,
'    formula_salary_code,
'    flag_ptkp,
'    ptkp_salary_code,
'    flag_pkp,
'    flag_pph21,
'    pph21_number,
'    flag_tax,
'    tax_salary_code,
'    flag_type,
'    flag_visible,
'    salary_value,
'    Description
'    )
'    SELECT
'    LEFT(dtm,7) AS MONTH,
'    str_employee_code AS employee_code,
'    salary_code,
'    str_company_code AS company_code,
'    salary_name,
'    dt1 AS date_from,
'    dt2 AS date_to,
'    flag_main_salary,
'    flag_sign,
'    flag_detail,
'    flag_use_formula,
'    formula_salary_code,
'    flag_ptkp,
'    ptkp_salary_code,
'    flag_pkp,
'    flag_pph21,
'    pph21_number,
'    0 AS flag_tax,
'    '' AS tax_salary_code,
'    'SU' AS flag_type,
'    flag_visible,
'    f_get_sum_dsum(str_employee_code,salary_code,dt1,dt2) AS salary_value,
'    Description
'    FROM m_salary_summary;

    strSQl = "select employee_code, employee_name from m_employee where department_code = '" & pDepartment & "'"
    rsemployee.Open strSQl, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsemployee.RecordCount > 0 Then
        rsemployee.MoveFirst
        While Not rsemployee.EOF
            strSQl = "insert into h_salary " & _
                "(MONTH, employee_code, salary_code, company_code, salary_name," & _
                "date_from, date_to, flag_main_salary, flag_sign,flag_detail," & _
                "flag_use_formula, formula_salary_code, flag_ptkp, ptkp_salary_code, flag_pkp," & _
                "flag_pph21, pph21_number, flag_tax, tax_salary_code, flag_type," & _
                "flag_visible, salary_value, Description)" & _
                "SELECT" & _
                    "(LEFT('" & pTgl1 & "',7) AS MONTH, '" & rsemployee!employee_code & "', salary_code," & _
                    "'" & COMPANY_CODE & "', salary_name, '" & pTgl1 & "', '" & pTgl2 & "', " & _
                    "flag_main_salary, flag_sign, flag_detail, flag_use_formula, " & _
                    "formula_salary_code, flag_ptkp, ptkp_salary_code, flag_pkp," & _
                    "flag_pph21, pph21_number, 0 AS flag_tax, '' AS tax_salary_code," & _
                    "'SU' AS flag_type, flag_visible," & _
                    "f_get_sum_dsum('" & rsemployee!employee_code & "',salary_code,'" & pTgl1 & "','" & pTgl2 & "'), Description) " & _
                "FROM m_salary_summary;"
            
            rsemployee.MoveNext
        Wend
    End If
End Sub

Public Sub insert_sallary_slip(department As String)

End Sub
