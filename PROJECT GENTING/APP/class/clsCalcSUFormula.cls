VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalcSUFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub CalcSuFormula(p_month As String, p_employee_code As String, _
        p_noJamsostek As String, p_npwp As String, p_Company_Code As String, _
        p_pph21_type As String, p_jstk_type As String, _
        p_strDate1 As String, p_strDate2 As String)

Dim v_salary_code, v_ptkp_salary_code, v_company_code As String
Dim v_flag_sign, v_flag_ptkp, v_flag_pkp, v_flag_pph21, v_pph21_number As Integer
Dim v_status, v_children As Integer
Dim v_salary, v_pph21, v_pph_value As Double
Dim vdt1, vdt2, vdt3 As String
Dim strsql As String
Dim rs_hsalary As New ADODB.Recordset
Dim rscari As New ADODB.Recordset
Dim NilaiHasil As Double
Dim v_npwp As String
Dim v_pph21_tanpa_npwp As Double
Dim v_sex, v_tanggung_pajak As Double
Dim v_pph21_code As String
Dim v_ptkp_code As String
Dim vTgl1 As String
Dim vTgl2 As String

Dim vBasicSalary As Double
Dim vFlagBPJSActive As Integer

Dim vMaxBPJSSalary As Double
Dim vMaxJPSalary As Double
Dim vSalary As Double

    strsql = "SELECT salary_code, flag_sign, flag_ptkp, flag_pkp, ptkp_salary_code, flag_pph21, pph21_number  " & _
        "FROM h_salary " & _
        "WHERE flag_type = 'SU' AND LEFT(MONTH,7) = '" & Left(p_month, 7) & "' " & _
        "AND employee_code = '" & p_employee_code & "' AND " & _
        "(flag_use_formula = 1 OR flag_ptkp = 1 OR flag_pkp = 1 OR flag_pph21 = 1) " & _
        "ORDER BY salary_code ASC"
    rs_hsalary.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs_hsalary.RecordCount > 0 Then
        rs_hsalary.MoveFirst
        While Not rs_hsalary.EOF
            v_salary_code = rs_hsalary!salary_code
            v_flag_sign = rs_hsalary!flag_sign
            v_flag_ptkp = rs_hsalary!flag_ptkp
            v_flag_pkp = rs_hsalary!flag_pkp
            v_ptkp_salary_code = rs_hsalary!ptkp_salary_code
            v_flag_pph21 = rs_hsalary!flag_pph21
            v_pph21_number = rs_hsalary!pph21_number
            
            vTgl1 = Format(frm_trans_salary_process.DTPicker_periode_from.Value, "yyyy-MM-dd")
            vTgl2 = Format(frm_trans_salary_process.DTPicker_periode_to.Value, "yyyy-MM-dd")
                
   
'            If v_salary_code = "SU-291" Then
'                strsql = "SELECT salary_value FROM h_salary " & _
'                        "WHERE employee_code = '" & p_employee_code & "' AND LEFT(MONTH,7) = '" & Left(p_month, 7) & "' " & _
'                        "AND salary_code = 'SU-06'"
'                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'                NilaiHasil = IIf(IsNull(rscari!salary_value), 0, rscari!salary_value) * 12
'                rscari.Close
'
'                If (0.05 * NilaiHasil) < 6000000 Then
'                   NilaiHasil = 0.05 * NilaiHasil
'                Else
'                   NilaiHasil = 6000000
'                End If
            If v_salary_code = "SU-286" Then
                strsql = "SELECT salary_value FROM h_salary " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND LEFT(MONTH,7) = '" & Left(p_month, 7) & "' " & _
                        "AND salary_code = 'SU-06'"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                NilaiHasil = IIf(IsNull(rscari!salary_value), 0, rscari!salary_value)
                rscari.Close
                                
                If (0.05 * NilaiHasil) < 500000 Then
                   NilaiHasil = 0.05 * NilaiHasil
                Else
                   NilaiHasil = 500000
                End If
                NilaiHasil = Round(NilaiHasil, 0)
                
'            ElseIf v_salary_code = "SU-30" Then
'                strSQL = "SELECT marital_status, number_of_children FROM m_employee WHERE employee_code = '" & p_employee_code & "'"
'                rscari.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
'                    v_status = rscari!marital_status
'                    v_children = IIf(IsNull(rscari!number_of_children), 0, rscari!number_of_children)
'                rscari.Close
'
'                strSQL = "SELECT f_get_ptkp('" & v_status & "', '" & v_children & "') ptkp_value"
'                rscari.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
'                    NilaiHasil = rscari!ptkp_value
'                rscari.Close

                               
            ElseIf v_flag_ptkp = 1 Then
                strsql = "SELECT marital_status, number_of_children, sex, pertanggungan_pajak FROM m_employee " & _
                        "WHERE employee_code = '" & p_employee_code & "'"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                    v_status = rscari!marital_status
                    v_children = IIf(IsNull(rscari!number_of_children), 0, rscari!number_of_children)
                    v_sex = rscari!sex
                    v_tanggung_pajak = IIf(IsNull(rscari!pertanggungan_pajak), 0, rscari!pertanggungan_pajak)
                rscari.Close
                
                strsql = "SELECT ptkp_type FROM m_salary_standard WHERE employee_code = '" & p_employee_code & "' " & _
                    "ORDER BY number DESC LIMIT 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                    If rscari.RecordCount > 0 Then
                        v_ptkp_code = rscari!ptkp_type
                    Else
                        v_ptkp_code = "STD"
                    End If
                rscari.Close
                
                strsql = "SELECT f_get_ptkp('" & v_status & "', '" & v_children & "', " & _
                        "'" & v_sex & "','" & v_tanggung_pajak & "','" & v_ptkp_code & "') ptkp_value"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                    NilaiHasil = rscari!ptkp_value
                    NilaiHasil = Round(NilaiHasil, 0)
                rscari.Close
'            /*
'            ElseIf int_flag_pkp = 1 Then
'               SET dbl_salary = (SELECT MAX(salary_value) FROM h_salary WHERE employee_code = str_employee_code
'                                 AND LEFT(MONTH,7)=LEFT(dt_month,7) AND salary_code = str_ptkp_salary_code);
'               SET dbl_salary = dbl_salary;
'            */
'            ElseIf v_flag_pph21 = 1 Then
'
'                strsql = "SELECT MAX(salary_value) pph_value FROM h_salary WHERE employee_code = '" & p_employee_code & "' " & _
'                    "AND LEFT(MONTH,7) = '" & Left(p_month, 7) & "' AND flag_pkp = 1"
'                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'                    v_pph_value = rscari!pph_value
'                rscari.Close
'
'                strsql = "SELECT pph21_type FROM m_salary_standard WHERE employee_code = '" & p_employee_code & "' " & _
'                    "ORDER BY number DESC LIMIT 1"
'                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'                    If rscari.RecordCount > 0 Then
'                        v_pph21_code = rscari!pph21_type
'                    Else
'                        v_pph21_code = "STD"
'                    End If
'                rscari.Close
'
'                strsql = "SELECT f_pph21(IFNULL('" & v_pph_value & "',0), '" & v_pph21_number & "', '" & v_pph21_code & "') pph21_value"
'                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'                    NilaiHasil = rscari!pph21_value
'                rscari.Close
            
'            ElseIf v_salary_code = "SU-075" Then
'                Dim tgl_awal, tgl_akhir As String
'                tgl_awal = p_month & "-01"
'                tgl_akhir = p_month & "-31"
'                strSQL = "SELECT fn_GetPotonganPinjam(p_employee_code,tgl_awal,tgl_akhir) potongan_pinjam"
'                rscari.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
'                If rscari.RecordCount > 0 Then
'                    NilaiHasil = rscari!potongan_pinjam
'                Else
'                    NilaiHasil = 0
'                End If
            
            Else
                strsql = "SELECT " & _
                  "SUM( " & _
                  "CASE WHEN formula_operation ='+' THEN ((f_get_sv_item('" & p_employee_code & "', b.formula_salary_code, '" & p_month & "') * formula_salary_sign) + (formula_variable * formula_variable_sign)) " & _
                       "WHEN formula_operation ='-' THEN ((f_get_sv_item('" & p_employee_code & "', b.formula_salary_code, '" & p_month & "') * formula_salary_sign) - (formula_variable * formula_variable_sign)) " & _
                       "WHEN formula_operation ='x' THEN ((f_get_sv_item('" & p_employee_code & "', b.formula_salary_code, '" & p_month & "') * formula_salary_sign) * (formula_variable * formula_variable_sign)) " & _
                       "WHEN formula_operation ='/' THEN ((f_get_sv_item('" & p_employee_code & "', b.formula_salary_code, '" & p_month & "') * formula_salary_sign) / (formula_variable * formula_variable_sign)) " & _
                  "END) * (+1) total " & _
                  "FROM h_salary a " & _
                  "JOIN m_salary_summary_formula b ON a.salary_code = b.salary_code " & _
                  "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = LEFT('" & p_month & "',7) AND a.salary_code = '" & v_salary_code & "'"
                  rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                    NilaiHasil = IIf(IsNull(rscari!total), 0, rscari!total)
                    NilaiHasil = Round(NilaiHasil, 0)
                rscari.Close
                
               If v_flag_pkp = 1 Then
                    NilaiHasil = IIf(NilaiHasil < 0, 0, NilaiHasil)
                    NilaiHasil = Round(NilaiHasil, 0)
               End If
            End If
            
            If v_salary_code = "SU-281" Then '--> POTONGAN JAMSOSTEK 2%
'                strsql = "SELECT b.salary_value, c.tk " & _
'                         "FROM m_employee a JOIN h_salary b ON a.employee_code = b.employee_code " & _
'                         "JOIN m_jamsostek_detail c ON a.company_code = c.company_code " & _
'                         "WHERE a.employee_code = '" & p_employee_code & "' " & _
'                            "AND LEFT(b.month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-01' " & _
'                            "AND c.jamsostek_code = '" & p_jstk_type & "'"
                strsql = "SELECT f_get_prorate('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10))  basic_salary, " & _
                              "f_get_functional('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) functional_allowance, " & _
                              "f_get_staff('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) staff_allowance, " & _
                              "f_get_skill('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) skill_allowance, " & _
                              "b.company_code, c.tk " & _
                         "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                         "JOIN m_jamsostek_detail c ON c.company_code = b.company_code " & _
                         "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " & _
                          "AND c.jamsostek_code = '" & p_jstk_type & "' AND date(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " & _
                         "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
'                        NilaiHasil = (rscari!tk / 100) * rscari!salary_value
                        NilaiHasil = (rscari!tk / 100) * (rscari!basic_salary + rscari!functional_allowance + rscari!staff_allowance + rscari!skill_allowance)
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
            
            ElseIf v_salary_code = "SU-282" Then
                'Mencari Nilai BPJS Kesehatan Karyawan
                If rscari.State Then rscari.Close
                strsql = "SELECT f_get_prorate('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10))  basic_salary, " & _
                              "f_get_functional('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) functional_allowance, " & _
                              "f_get_staff('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) staff_allowance, " & _
                              "f_get_skill('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) skill_allowance, " & _
                        "b.company_code, c.kary_value, c.max_bpjs_salary, " & _
                        "(SELECT no_bpjs FROM t_bpjs WHERE employee_code = a.employee_code AND DATE(reg_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' ORDER BY reg_date DESC LIMIT 1) no_bpjs," & _
                        "(SELECT flag_active FROM t_bpjs WHERE employee_code = a.employee_code AND DATE(reg_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' ORDER BY reg_date DESC LIMIT 1) flag_active " & _
                      "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                      "JOIN m_bpjs c ON c.company_code = b.company_code " & _
                      "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " & _
                        "AND DATE(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " & _
                      "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vMaxBPJSSalary = rscari!max_bpjs_salary
                    
                    vBasicSalary = rscari!basic_salary
                    vFlagBPJSActive = IIf(IsNull(rscari!flag_active), 0, rscari!flag_active)
                    
                    If vFlagBPJSActive = 1 Then
                        If vBasicSalary <> 0 Then
                            If IIf(IsNull(rscari!no_bpjs), "", rscari!no_bpjs) <> "" Then
                                vSalary = (rscari!basic_salary + rscari!functional_allowance + rscari!staff_allowance + rscari!skill_allowance)
                        
                                If vSalary > vMaxBPJSSalary Then
                                    NilaiHasil = (rscari!kary_value / 100) * vMaxBPJSSalary
                                Else
                                    NilaiHasil = (rscari!kary_value / 100) * vSalary
                                End If
                            Else
                                NilaiHasil = 0
                            End If
                        Else
                            NilaiHasil = 0
                        End If
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
              
            ElseIf v_salary_code = "SU-283" Then
                Dim vBasic As Double
                'Mencari Nilai JP
                If rscari.State Then rscari.Close
                strsql = "SELECT f_get_prorate('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10))  basic_salary, " & _
                              "f_get_functional('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) functional_allowance, " & _
                              "f_get_staff('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) staff_allowance, " & _
                              "f_get_skill('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) skill_allowance, " & _
                        "b.company_code, c.jp_emp, c.max_jp_salary " & _
                      "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                      "JOIN m_jamsostek_detail c ON c.company_code = b.company_code " & _
                      "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " & _
                        "AND c.jamsostek_code = '" & p_jstk_type & "' AND date(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " & _
                      "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vMaxJPSalary = rscari!max_jp_salary
                    
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
                        vSalary = (rscari!basic_salary + rscari!functional_allowance + rscari!staff_allowance + rscari!skill_allowance)
                        
                        If vSalary > vMaxJPSalary Then
                            NilaiHasil = (rscari!jp_emp / 100) * vMaxJPSalary
                        Else
                            NilaiHasil = (rscari!jp_emp / 100) * vSalary
                        End If
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
              
            ElseIf v_salary_code = "SU-36A" Then '--> POTONGAN PPh TANPA NPWP
                strsql = "SELECT salary_value " & _
                        "FROM h_salary " & _
                        "WHERE employee_code = '" & p_employee_code & "' " & _
                        "AND LEFT(month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-36'"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

                v_pph21_tanpa_npwp = rscari!salary_value
                v_npwp = p_npwp

                If rscari.RecordCount > 0 Then
'                    If IIf(IsNull(v_npwp) Or v_npwp = "" Or v_npwp = "-", "0", v_npwp) = "0" Then
                    If IIf(IsNull(v_npwp), "", v_npwp) = "" Then
                        NilaiHasil = 0.2 * v_pph21_tanpa_npwp
                    Else
                        NilaiHasil = 0
                    End If
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
                
            ElseIf v_salary_code = "SU-285" Then '--> POTONGAN JSTK
'                strsql = "SELECT b.salary_value, c.tk " & _
'                         "FROM m_employee a JOIN h_salary b ON a.employee_code = b.employee_code " & _
'                         "JOIN m_jamsostek_detail c ON a.company_code = c.company_code " & _
'                         "WHERE a.employee_code = '" & p_employee_code & "' " & _
'                            "AND LEFT(b.month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-01' " & _
'                            "AND c.jamsostek_code = '" & p_jstk_type & "'"
                
                strsql = "SELECT f_get_prorate('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10))  basic_salary, " & _
                              "f_get_functional('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) functional_allowance, " & _
                              "f_get_staff('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) staff_allowance, " & _
                              "f_get_skill('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) skill_allowance, " & _
                              "b.company_code, c.tk " & _
                         "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                         "JOIN m_jamsostek_detail c ON c.company_code = b.company_code " & _
                         "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " & _
                          "AND c.jamsostek_code = '" & p_jstk_type & "' AND date(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " & _
                         "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

                If rscari.RecordCount > 0 Then
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
'                        NilaiHasil = (rscari!tk / 100) * rscari!salary_value
                        NilaiHasil = (rscari!tk / 100) * (rscari!basic_salary + rscari!functional_allowance + rscari!staff_allowance + rscari!skill_allowance)
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
'                End If
                rscari.Close
            
            ElseIf v_salary_code = "SU-36" Then '--> POTONGAN PPh Sebulan
                strsql = "SELECT a.start_working, b.salary_value, b.employee_code " & _
                        "FROM m_employee a join h_salary b on a.employee_code = b.employee_code " & _
                        "WHERE a.employee_code = '" & p_employee_code & "' " & _
                        "AND LEFT(b.month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-35A'"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
'                    If ((DateDiff("m", rscari!start_working, p_month)) + 1) < 12 And year(rscari!start_working) = year(p_month) Then
                    If ((DateDiff("m", rscari!start_working, year(Now) & "-12-31")) + 1) < 12 And year(rscari!start_working) = year(p_month) Then
                        If rscari!salary_value = 0 Then
                            NilaiHasil = 0
                        Else
                            NilaiHasil = rscari!salary_value / ((DateDiff("m", rscari!start_working, year(Now) & "-12-31")) + 1)
                        End If
                    Else
                        If rscari!salary_value = 0 Then
                            NilaiHasil = 0
                        Else
'                            If rscari!employee_code = 15 Then
'                                NilaiHasil = (rscari!salary_value - 6739000) / 12
'                            Else
                                NilaiHasil = rscari!salary_value / 12
'                            End If
                        End If
                    End If
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
                
            ElseIf v_salary_code = "SU-2891" Then '--> Annual Salary
                strsql = "SELECT a.start_working, b.salary_value " & _
                        "FROM m_employee a join h_salary b on a.employee_code = b.employee_code " & _
                        "WHERE a.employee_code = '" & p_employee_code & "' " & _
                        "AND LEFT(b.month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-289'"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    a = (DateDiff("m", rscari!start_working, year(Now) & "-12-31")) + 1
                    If a < 12 And year(rscari!start_working) = year(p_month) Then
'                        NilaiHasil = rscari!salary_value * ((DateDiff("m", rscari!start_working, p_month)) + 1)
                        NilaiHasil = rscari!salary_value * ((DateDiff("m", rscari!start_working, year(Now) & "-12-31")) + 1)
                    Else
                        NilaiHasil = rscari!salary_value * 12
                    End If
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
                
            ElseIf v_salary_code = "SU-0591" Then 'JKK 0.24% or 0.54%

'                strsql = "SELECT b.salary_value, c.jkk " & _
'                         "FROM m_employee a JOIN h_salary b ON a.employee_code = b.employee_code " & _
'                         "JOIN m_jamsostek_detail c ON a.company_code = c.company_code " & _
'                         "WHERE a.employee_code = '" & p_employee_code & "' " & _
'                            "AND LEFT(b.month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-01' " & _
'                            "AND c.jamsostek_code = '" & p_jstk_type & "'"
                
                strsql = "SELECT f_get_prorate('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10))  basic_salary, " & _
                              "f_get_functional('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) functional_allowance, " & _
                              "f_get_staff('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) staff_allowance, " & _
                              "f_get_skill('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) skill_allowance, " & _
                              "b.company_code, c.jkk " & _
                         "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                         "JOIN m_jamsostek_detail c ON c.company_code = b.company_code " & _
                         "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " & _
                          "AND c.jamsostek_code = '" & p_jstk_type & "' AND date(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " & _
                         "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
'                        NilaiHasil = (rscari!jkk / 100) * rscari!salary_value
                        NilaiHasil = (rscari!jkk / 100) * (rscari!basic_salary + rscari!functional_allowance + rscari!staff_allowance + rscari!skill_allowance)
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
            
            ElseIf v_salary_code = "SU-0592" Then 'JKM 0.3%
'                strsql = "SELECT b.salary_value, c.jkm " & _
'                         "FROM m_employee a JOIN h_salary b ON a.employee_code = b.employee_code " & _
'                         "JOIN m_jamsostek_detail c ON a.company_code = c.company_code " & _
'                         "WHERE a.employee_code = '" & p_employee_code & "' " & _
'                            "AND LEFT(b.month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-01' " & _
'                            "AND c.jamsostek_code = '" & p_jstk_type & "'"
                
                strsql = "SELECT f_get_prorate('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10))  basic_salary, " & _
                              "f_get_functional('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) functional_allowance, " & _
                              "f_get_staff('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) staff_allowance, " & _
                              "f_get_skill('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) skill_allowance, " & _
                              "b.company_code, c.jkm " & _
                         "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                         "JOIN m_jamsostek_detail c ON c.company_code = b.company_code " & _
                         "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " & _
                          "AND c.jamsostek_code = '" & p_jstk_type & "' AND date(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " & _
                         "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
'                        NilaiHasil = (rscari!jkm / 100) * rscari!salary_value
                        NilaiHasil = (rscari!jkm / 100) * (rscari!basic_salary + rscari!functional_allowance + rscari!staff_allowance + rscari!skill_allowance)
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
            ElseIf v_salary_code = "SU-0593" Then 'Jamsostek Perusahaan 3.7%
'                strsql = "SELECT b.salary_value, c.prs " & _
'                         "FROM m_employee a JOIN h_salary b ON a.employee_code = b.employee_code " & _
'                         "JOIN m_jamsostek_detail c ON a.company_code = c.company_code " & _
'                         "WHERE a.employee_code = '" & p_employee_code & "' " & _
'                            "AND LEFT(b.month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-01' " & _
'                            "AND c.jamsostek_code = '" & p_jstk_type & "'"
                
                strsql = "SELECT f_get_prorate('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10))  basic_salary, " & _
                              "f_get_functional('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) functional_allowance, " & _
                              "f_get_staff('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) staff_allowance, " & _
                              "f_get_skill('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) skill_allowance, " & _
                              "b.company_code, c.prs " & _
                         "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                         "JOIN m_jamsostek_detail c ON c.company_code = b.company_code " & _
                         "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " & _
                          "AND c.jamsostek_code = '" & p_jstk_type & "' AND date(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " & _
                         "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
'                        NilaiHasil = (rscari!prs / 100) * rscari!salary_value
                        NilaiHasil = (rscari!prs / 100) * (rscari!basic_salary + rscari!functional_allowance + rscari!staff_allowance + rscari!skill_allowance)
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
            
            ElseIf v_salary_code = "SU-0594" Then
                'Mencari Nilai BPJS Kesehatan Perusahaan
                If rscari.State Then rscari.Close
                SQL = "SELECT f_get_prorate('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10))  basic_salary, " & _
                              "f_get_functional('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) functional_allowance, " & _
                              "f_get_staff('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) staff_allowance, " & _
                              "f_get_skill('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) skill_allowance, " & _
                        "b.company_code, c.pers_value, c.max_bpjs_salary, " & _
                        "(SELECT no_bpjs FROM t_bpjs WHERE employee_code = a.employee_code AND DATE(reg_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' ORDER BY reg_date DESC LIMIT 1) no_bpjs," & _
                        "(SELECT flag_active FROM t_bpjs WHERE employee_code = a.employee_code AND DATE(reg_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' ORDER BY reg_date DESC LIMIT 1) flag_active " & _
                      "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                      "JOIN m_bpjs c ON c.company_code = b.company_code " & _
                      "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " & _
                        "AND DATE(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " & _
                      "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    vMaxBPJSSalary = rscari!max_bpjs_salary
                    
                    vBasicSalary = rscari!basic_salary
                    vFlagBPJSActive = IIf(IsNull(rscari!flag_active), 0, rscari!flag_active)
                    
                    If vFlagBPJSActive = 1 Then
                        If vBasicSalary <> 0 Then
                            If IIf(IsNull(rscari!no_bpjs), "", rscari!no_bpjs) <> "" Then
                                vSalary = (rscari!basic_salary + rscari!functional_allowance + rscari!staff_allowance + rscari!skill_allowance)
                        
                                If vSalary > vMaxBPJSSalary Then
                                    NilaiHasil = (rscari!pers_value / 100) * vMaxBPJSSalary
                                Else
                                    NilaiHasil = (rscari!pers_value / 100) * vSalary
                                End If
                            Else
                                NilaiHasil = 0
                            End If
                        Else
                            NilaiHasil = 0
                        End If
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
            ElseIf v_salary_code = "SU-0595" Then
                'Mencari Nilai JP Perusahaan
                If rscari.State Then rscari.Close
                strsql = "SELECT f_get_prorate('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10))  basic_salary, " & _
                              "f_get_functional('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) functional_allowance, " & _
                              "f_get_staff('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) staff_allowance, " & _
                              "f_get_skill('" & p_employee_code & "', LEFT('" & p_strDate1 & "',10), LEFT('" & p_strDate2 & "',10)) skill_allowance, " & _
                        "b.company_code, c.jp_co, c.max_jp_salary " & _
                      "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                      "JOIN m_jamsostek_detail c ON c.company_code = b.company_code " & _
                      "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " & _
                        "AND c.jamsostek_code = '" & p_jstk_type & "' AND date(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " & _
                      "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    vMaxJPSalary = rscari!max_jp_salary
                    
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
                        vSalary = (rscari!basic_salary + rscari!functional_allowance + rscari!staff_allowance + rscari!skill_allowance)
                        
                        If vSalary > vMaxJPSalary Then
                            NilaiHasil = (rscari!jp_co / 100) * vMaxJPSalary
                        Else
                            NilaiHasil = (rscari!jp_co / 100) * vSalary
                        End If
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
            '+++++++++++++++++++++++++++++++ THR ++++++++++++++++
            ElseIf v_salary_code = "SU-052" Then
                strsql = "SELECT jmlthr " & _
                        "from t_thr " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND year(tgltrans) = '" & Left(p_month, 4) & "' " & _
                        "AND month(tgltrans) = '" & Mid(p_month, 6, 2) & "' AND jenis = 0"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = rscari!jmlTHR
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            '++++++++++++++++++++++++++++++++++++++++
            
            '+++++++++++++++++++++++++++++++ Bonus ++++++++++++++++
            ElseIf v_salary_code = "SU-054" Then
                strsql = "SELECT jmlthr " & _
                        "from t_thr " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND year(tgltrans) = '" & Left(p_month, 4) & "' " & _
                        "AND month(tgltrans) = '" & Mid(p_month, 6, 2) & "' AND jenis = 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = rscari!jmlTHR
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            '++++++++++++++++++++++++++++++++++++++++
            
            '+++++++++++++++++++++++++++++++ Rapel ++++++++++++++++
            ElseIf v_salary_code = "SU-0541" Then
                strsql = "SELECT SUM(jmlPotong) jmlPotong " & _
                        "from t_employee_income " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND flag_other_income = 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = IIf(IsNull(rscari!jmlpotong), 0, rscari!jmlpotong)
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            '++++++++++++++++++++++++++++++++++++++++
            
            '+++++++++++++++++++++++++++++++ Returnable Ammount ++++++++++++++++
            ElseIf v_salary_code = "SU-0542" Then
                strsql = "SELECT SUM(jmlPotong) jmlPotong " & _
                        "from t_employee_income " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND flag_other_income = 0"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = IIf(IsNull(rscari!jmlpotong), 0, rscari!jmlpotong)
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            '++++++++++++++++++++++++++++++++++++++++
            
            '+++++++++++++++++++++++++++++++ Lembur Temporary ++++++++++++++++
            ElseIf v_salary_code = "SU-0543" Then
                strsql = "SELECT SUM(jmlPotong) jmlPotong " & _
                        "from t_employee_income " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND flag_other_income = 4"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = IIf(IsNull(rscari!jmlpotong), 0, rscari!jmlpotong)
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            '++++++++++++++++++++++++++++++++++++++++
            
            '+++++++++++++++++++++++++++++++ Misc. Adjusment In ++++++++++++++++
            ElseIf v_salary_code = "SU-0544" Then
                strsql = "SELECT SUM(jmlPotong) jmlPotong " & _
                        "from t_employee_income " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND flag_other_income = 2"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = IIf(IsNull(rscari!jmlpotong), 0, rscari!jmlpotong)
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            '++++++++++++++++++++++++++++++++++++++++
            
            '+++++++++++++++++++++++++++++++ Other Allowance ++++++++++++++++
            ElseIf v_salary_code = "SU-059" Then
                strsql = "SELECT SUM(jmlPotong) jmlPotong " & _
                        "from t_employee_income " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND flag_other_income = 3"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = IIf(IsNull(rscari!jmlpotong), 0, rscari!jmlpotong)
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            '++++++++++++++++++++++++++++++++++++++++
            
'            '+++++++++++++++++++++++++++++++ Ekspense ++++++++++++++++
'            ElseIf v_salary_code = "SU-073" Then
'                strsql = "SELECT jmlPotong " &
'                        "from t_employee_expense " &
'                        "WHERE employee_code = '" & p_employee_code & "' AND year(tgltrans) = '" & Left(p_month, 4) & "' " &
'                        "AND month(tgltrans) = '" & Mid(p_month, 6, 2) & "'"
'                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'                If rscari.RecordCount > 0 Then
'                    NilaiHasil = rscari!jmlpotong
'                Else
'                    NilaiHasil = 0
'                End If
'                rscari.Close
'            '++++++++++++++++++++++++++++++++++++++++
                
            '+++++++++++++++++++++++++++++++ Tax Shortage ++++++++++++++++
            ElseIf v_salary_code = "SU-0711" Then
                strsql = "SELECT SUM(jmlPotong) jmlPotong " & _
                        "from t_employee_expense " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND flag_other_expense = 0"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = IIf(IsNull(rscari!jmlpotong), 0, rscari!jmlpotong)
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            '++++++++++++++++++++++++++++++++++++++++
            
            '+++++++++++++++++++++++++++++++ Miscellaneous ++++++++++++++++
            ElseIf v_salary_code = "SU-0712" Then
'                strsql = "SELECT SUM(jmlPotong) jmlPotong " & _
'                        "from t_employee_expense " & _
'                        "WHERE employee_code = '" & p_employee_code & "' AND year(tgltrans) = '" & Left(p_month, 4) & "' " & _
'                        "AND month(tgltrans) = '" & Mid(p_month, 6, 2) & "' AND flag_other_expense = 1"
                
                strsql = "SELECT SUM(jmlPotong) jmlPotong " & _
                        "from t_employee_expense " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND flag_other_expense = 1"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = IIf(IsNull(rscari!jmlpotong), 0, rscari!jmlpotong)
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            '++++++++++++++++++++++++++++++++++++++++
            
            '+++++++++++++++++++++++++++++++ Misc. Deduction In ++++++++++++++++
            ElseIf v_salary_code = "SU-0713" Then
                strsql = "SELECT SUM(jmlPotong) jmlPotong " & _
                        "from t_employee_expense " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND flag_other_expense = 2"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = IIf(IsNull(rscari!jmlpotong), 0, rscari!jmlpotong)
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            '++++++++++++++++++++++++++++++++++++++++
            
'            ElseIf v_salary_code = "SU-293" Then 'JAMSOSTEK LIFE 0.3%
'                strsql = "SELECT main_salary,jstk_life_allowance " & _
'                        "from m_salary_standard " & _
'                        "WHERE employee_code = '" & p_employee_code & "'"
'                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'                If rscari.RecordCount > 0 Then
'                    NilaiHasil = (rscari!jstk_life_allowance / 100) * rscari!main_salary
'                Else
'                    NilaiHasil = 0
'                End If
'
'                rscari.Close
            
            '++++++++++++++++++++++++++++++++++++ PPh21 ++++++++++++++++++++++++++++++++
            ElseIf v_salary_code = "SU-32" Then 'PPh 5%
                
                strsql = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 1 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-31'"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                
                    If rscari!salary_value < 0 Then
                        NilaiHasil = 0
                    Else
                        If rscari!salary_value > rscari!pph21_upper Then '50000000
                            NilaiHasil = (rscari!pph21_percentage / 100) * rscari!pph21_upper '50000000
                        Else
                            NilaiHasil = (rscari!pph21_percentage / 100) * rscari!salary_value
                        End If
                    End If
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            ElseIf v_salary_code = "SU-33" Then 'PPh21 15%
            
                strsql = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 2 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-31'"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    If rscari!salary_value < 0 Then
                        NilaiHasil = 0
                    Else
                        If rscari!salary_value <= rscari!pph21_under Then '50000000
                            NilaiHasil = 0
                        ElseIf rscari!salary_value < rscari!pph21_upper Then '250000000
                            NilaiHasil = (rscari!pph21_percentage / 100) * (rscari!salary_value - rscari!pph21_under) '50000000)
                        Else
                            NilaiHasil = (rscari!pph21_percentage / 100) * (rscari!pph21_upper - rscari!pph21_under) '200000000
                        End If
                    End If
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            ElseIf v_salary_code = "SU-34" Then 'PPh21 25%
            
                strsql = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 3 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-31'"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    If rscari!salary_value < 0 Then
                        NilaiHasil = 0
                    Else
                        If rscari!salary_value <= rscari!pph21_under Then '50000000
                            NilaiHasil = 0
                        ElseIf rscari!salary_value < rscari!pph21_upper Then '250000000
                            NilaiHasil = (rscari!pph21_percentage / 100) * (rscari!salary_value - rscari!pph21_under) '50000000)
                        Else
                            NilaiHasil = (rscari!pph21_percentage / 100) * (rscari!pph21_upper - rscari!pph21_under) '200000000
                        End If
                    End If
                End If
                rscari.Close
                
                NilaiHasil = Round(NilaiHasil, 0)
            ElseIf v_salary_code = "SU-35" Then 'PPh21 30%
            
                strsql = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 4 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-31'"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    If rscari!salary_value < 0 Then
                        NilaiHasil = 0
                    Else
                        If rscari!salary_value <= rscari!pph21_under Then '500000000
                            NilaiHasil = 0
                        Else
                            If rscari!salary_value > rscari!pph21_under Then '500000000
                                NilaiHasil = (rscari!pph21_percentage / 100) * (rscari!salary_value - rscari!pph21_under) '500000000)
                            Else
                                NilaiHasil = 0
                            End If
                        End If
                    End If
                End If
                rscari.Close
                NilaiHasil = Round(NilaiHasil, 0)
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            ElseIf v_salary_code = "SU-77" Then
                strsql = "SELECT salary_value FROM h_salary " & _
                  "WHERE employee_code = '" & p_employee_code & "' " & _
                  "AND LEFT(MONTH,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-37'"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                    NilaiHasil = IIf(IsNull(NilaiHasil), 0, NilaiHasil)
                rscari.Close
                
                strsql = "SELECT s_date1,s_date2,s_date3 INTO vdt1,vdt2,vdt3 FROM s_est WHERE s_number=1"
                CnG.Execute strsql
                
                strsql = "Select (f_count_presence(str_employee_code, vdt1, vdt3)" & _
                    "/(f_count_working_days(str_company_code, str_employee_code, vdt1, vdt2))" & _
                    "* dbl_salary) total"
                rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                    NilaiHasil = rscari!total
                    NilaiHasil = Round(NilaiHasil, 0)
                rscari.Close
            End If
            
            '+++++++++++++++++++++++++++++++++++++++++++++
            strsql = "UPDATE h_salary SET salary_value = '" & NilaiHasil & "' " & _
                  "WHERE LEFT(MONTH,7) = LEFT('" & p_month & "',7) " & _
                  "AND employee_code = '" & p_employee_code & "' AND salary_code = '" & v_salary_code & "'"
            CnG.Execute strsql
            '+++++++++++++++++++++++++++++++++++++++++++++
                        
            rs_hsalary.MoveNext
        Wend
    End If
End Sub
