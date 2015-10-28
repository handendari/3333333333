VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strsql As String

Public Sub InsertLog(remark As String)
    strsql = "INSERT log_activitas (user_code,login_time,click_time,description,company_code,department_code) " & _
        "VALUES " & _
        "('" & LOGIN_CODE & "','" & Format(loginTime, "yyyy-MM-dd hh:mm:ss") & "',now(),'" & remark & "','" & COMPANY_CODE & "','" & DEPARTMENT_CODE & "')"
        
    CnG.Execute strsql
End Sub

Public Function isSalaryStandardExist(kdEmployee As String) As Boolean
Dim rsSalaryExist As New ADODB.Recordset
    strsql = "SELECT 1 FROM m_salary_standard WHERE employee_code = '" & kdEmployee & "'"
    rsSalaryExist.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsSalaryExist.RecordCount > 0 Then
        isSalaryStandardExist = True
    Else
        isSalaryStandardExist = False
    End If
    rsSalaryExist.Close
End Function

Public Function getUserLevel(kdEmployee As String) As Integer
Dim rsUserLevel As New ADODB.Recordset
    strsql = "SELECT ifnull(level_code,0) level_code FROM m_employee WHERE employee_code = '" & kdEmployee & "'"
    rsUserLevel.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsUserLevel.RecordCount > 0 Then
        getUserLevel = IIf(rsUserLevel!level_code = 0, 100, rsUserLevel!level_code)
    End If
    rsUserLevel.Close
End Function

Public Function isEmployeeExist(kdEmployee As String) As Boolean
Dim rscari As New ADODB.Recordset
    strsql = "SELECT 1 FROM m_employee WHERE employee_code = '" & kdEmployee & "'"
    rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        isEmployeeExist = True
    End If
    rscari.Close
End Function

