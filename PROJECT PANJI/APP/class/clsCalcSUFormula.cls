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
        p_pph21_type As String, p_jstk_type As String)

Dim v_salary_code, v_ptkp_salary_code, v_company_code As String
Dim v_flag_sign, v_flag_ptkp, v_flag_pkp, v_flag_pph21, v_pph21_number As Integer
Dim v_status, v_children As Integer
Dim v_salary, v_pph21, v_pph_value As Double
Dim vdt1, vdt2, vdt3 As String
Dim SQL As String
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
Dim vUMK As Double

    SQL = "SELECT salary_code, flag_sign, flag_ptkp, flag_pkp, ptkp_salary_code, flag_pph21, pph21_number  " & _
        "FROM h_salary " & _
        "WHERE flag_type = 'SU' AND LEFT(MONTH,7) = '" & Left(p_month, 7) & "' " & _
        "AND employee_code = '" & p_employee_code & "' AND " & _
        "(flag_use_formula = 1 OR flag_ptkp = 1 OR flag_pkp = 1 OR flag_pph21 = 1) " & _
        "ORDER BY salary_code ASC"
    rs_hsalary.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
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
   
            If v_salary_code = "SU-286" Then 'Biaya Jabatan
                Dim vPersenBiayaJabatan As Double
                Dim vMaksBiayaJabatan As Double
                Dim vMaksJbtnSebulan As Double
                Dim vPersenJabatan As Double
                
                SQL = "SELECT salary_value FROM h_salary " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND LEFT(MONTH,7) = '" & Left(p_month, 7) & "' " & _
                        "AND salary_code = 'SU-06'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                NilaiHasil = IIf(IsNull(rscari!salary_value), 0, rscari!salary_value)
                rscari.Close
                                
                SQL = "SELECT prsn_jbtn, maks_jbtn FROM m_pph21 WHERE pph21_code = '" & p_pph21_type & "'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vPersenBiayaJabatan = IIf(IsNull(rscari!prsn_jbtn), 0, rscari!prsn_jbtn)
                    vMaksBiayaJabatan = IIf(IsNull(rscari!maks_jbtn), 0, rscari!maks_jbtn)
                End If
                rscari.Close
                
                vMaksJbtnSebulan = vMaksBiayaJabatan / 12
                vPersenJabatan = vPersenBiayaJabatan / 100
                
                If (vPersenJabatan * NilaiHasil) < vMaksJbtnSebulan Then
                   NilaiHasil = (vPersenBiayaJabatan / 100) * NilaiHasil
                Else
                   NilaiHasil = vMaksJbtnSebulan
                End If
                               
            ElseIf v_flag_ptkp = 1 Then
                SQL = "SELECT marital_status, no_of_children, sex FROM m_employee " & _
                        "WHERE employee_code = '" & p_employee_code & "'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    v_status = rscari!marital_status
                    v_children = IIf(IsNull(rscari!no_of_children), 0, rscari!no_of_children)
                    v_sex = rscari!sex
                rscari.Close
                
                SQL = "SELECT ptkp_type FROM m_salary_standard WHERE employee_code = '" & p_employee_code & "' AND date(salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " & _
                    "ORDER BY salary_date DESC LIMIT 1"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    If rscari.RecordCount > 0 Then
                        v_ptkp_code = rscari!ptkp_type
                    Else
                        v_ptkp_code = "STD"
                    End If
                rscari.Close
                
                SQL = "SELECT f_get_ptkp('" & v_status & "', '" & v_children & "', " & _
                        "'" & v_sex & "','" & v_tanggung_pajak & "','" & v_ptkp_code & "') ptkp_value"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    NilaiHasil = rscari!ptkp_value
                rscari.Close
            
            Else
                SQL = "SELECT " & _
                  "SUM( " & _
                  "CASE WHEN formula_operation ='+' THEN ((f_get_sv_item('" & p_employee_code & "', b.formula_salary_code, '" & p_month & "') * formula_salary_sign) + (formula_variable * formula_variable_sign)) " & _
                       "WHEN formula_operation ='-' THEN ((f_get_sv_item('" & p_employee_code & "', b.formula_salary_code, '" & p_month & "') * formula_salary_sign) - (formula_variable * formula_variable_sign)) " & _
                       "WHEN formula_operation ='x' THEN ((f_get_sv_item('" & p_employee_code & "', b.formula_salary_code, '" & p_month & "') * formula_salary_sign) * (formula_variable * formula_variable_sign)) " & _
                       "WHEN formula_operation ='/' THEN ((f_get_sv_item('" & p_employee_code & "', b.formula_salary_code, '" & p_month & "') * formula_salary_sign) / (formula_variable * formula_variable_sign)) " & _
                  "END) * (+1) total " & _
                  "FROM h_salary a " & _
                  "JOIN m_salary_summary_formula b ON a.salary_code = b.salary_code " & _
                  "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = LEFT('" & p_month & "',7) AND a.salary_code = '" & v_salary_code & "'"
                  rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    NilaiHasil = IIf(IsNull(rscari!total), 0, rscari!total)
                rscari.Close
                
               If v_flag_pkp = 1 Then
                    NilaiHasil = IIf(NilaiHasil < 0, 0, NilaiHasil)
               End If
            End If
            
            If v_salary_code = "SU-281" Then '--> POTONGAN JAMSOSTEK 2%
'                SQL = "SELECT a.salary_value,c.tk " & _
'                        "FROM h_salary a join m_employee b on a.employee_code = b.employee_code " & _
'                            "JOIN m_jamsostek_detail c on b.company_code = c.company_code " & _
'                        "WHERE a.employee_code = '" & p_employee_code & "' " & _
'                        "AND LEFT(month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-01'"
                
                SQL = "SELECT umk_value FROM m_pref_umk WHERE date(umk_date) <= '" & Format(frm_trans_salary_process.DTPicker_periode_to.Value, "yyyy-MM-dd") & "' " & _
                        "ORDER by umk_date DESC limit 1"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vUMK = rscari!umk_value
                Else
                    vUMK = 0
                End If
                rscari.Close
                
                SQL = "SELECT a.salary_value,c.tk " & _
                        "FROM h_salary a join m_employee b on a.employee_code = b.employee_code " & _
                            "JOIN m_jamsostek_detail c on b.company_code = c.company_code " & _
                        "WHERE a.employee_code = '" & p_employee_code & "' " & _
                        "AND LEFT(month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-01'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
                        NilaiHasil = (rscari!tk / 100) * vUMK
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
            
            ElseIf v_salary_code = "SU-36A" Then '--> POTONGAN PPh TANPA NPWP
                Dim vTrfNonNPWP As Double
                
                SQL = "SELECT salary_value " & _
                        "FROM h_salary " & _
                        "WHERE employee_code = '" & p_employee_code & "' " & _
                        "AND LEFT(month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-36'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    v_pph21_tanpa_npwp = rscari!salary_value
                Else
                    v_pph21_tanpa_npwp = 0
                End If
                v_npwp = IIf(IsNull(p_npwp), "", p_npwp)
                rscari.Close
                
                SQL = "SELECT trf_non_npwp FROM m_pph21 WHERE pph21_code = '" & p_pph21_type & "'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vTrfNonNPWP = IIf(IsNull(rscari!trf_non_npwp), 0, rscari!trf_non_npwp)
                Else
                    vTrfNonNPWP = 0
                End If
                rscari.Close
                
                If IIf(IsNull(v_npwp), "", v_npwp) = "" Then
                    NilaiHasil = (vTrfNonNPWP / 100) * v_pph21_tanpa_npwp
                Else
                    NilaiHasil = 0
                End If

            ElseIf v_salary_code = "SU-285" Then '--> POTONGAN JSTK
                SQL = "SELECT umk_value FROM m_pref_umk WHERE date(umk_date) <= '" & Format(frm_trans_salary_process.DTPicker_periode_to.Value, "yyyy-MM-dd") & "' " & _
                        "ORDER by umk_date DESC limit 1"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vUMK = rscari!umk_value
                Else
                    vUMK = 0
                End If
                rscari.Close
                
                SQL = "SELECT a.salary_value,c.tk " & _
                        "FROM h_salary a join m_employee b on a.employee_code = b.employee_code " & _
                            "JOIN m_jamsostek_detail c on b.company_code = c.company_code " & _
                        "WHERE a.employee_code = '" & p_employee_code & "' " & _
                        "AND LEFT(month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-01'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
                        NilaiHasil = (rscari!tk / 100) * vUMK
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
            
            ElseIf v_salary_code = "SU-36" Then '--> POTONGAN PPh Sebulan
                SQL = "SELECT a.start_working, b.salary_value, b.employee_code " & _
                        "FROM m_employee a join h_salary b on a.employee_code = b.employee_code " & _
                        "WHERE a.employee_code = '" & p_employee_code & "' " & _
                        "AND LEFT(b.month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-35A'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
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
                            NilaiHasil = rscari!salary_value / 12
                        End If
                    End If
                End If
                rscari.Close
            
            ElseIf v_salary_code = "SU-361" Then '--> POTONGAN PPh Sebulan - GROSS UP
                SQL = "SELECT a.start_working, b.salary_value, b.employee_code " & _
                        "FROM m_employee a join h_salary b on a.employee_code = b.employee_code " & _
                        "WHERE a.employee_code = '" & p_employee_code & "' " & _
                        "AND LEFT(b.month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-35B'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
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
                            NilaiHasil = rscari!salary_value / 12
                        End If
                    End If
                End If
                rscari.Close
                
            ElseIf v_salary_code = "SU-2891" Then '--> Annual Salary
                SQL = "SELECT a.start_working, b.salary_value " & _
                        "FROM m_employee a join h_salary b on a.employee_code = b.employee_code " & _
                        "WHERE a.employee_code = '" & p_employee_code & "' " & _
                        "AND LEFT(b.month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-289'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    a = (DateDiff("m", rscari!start_working, year(Now) & "-12-31")) + 1
                    If a < 12 And year(rscari!start_working) = year(p_month) Then
                        NilaiHasil = rscari!salary_value * ((DateDiff("m", rscari!start_working, year(Now) & "-12-31")) + 1)
                    Else
                        NilaiHasil = rscari!salary_value * 12
                    End If
                End If
                rscari.Close
            
            ElseIf v_salary_code = "SU-2892" Then '--> Annual Salary - GROSS UP
                SQL = "SELECT a.start_working, b.salary_value " & _
                        "FROM m_employee a join h_salary b on a.employee_code = b.employee_code " & _
                        "WHERE a.employee_code = '" & p_employee_code & "' " & _
                        "AND LEFT(b.month,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-2890'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    a = (DateDiff("m", rscari!start_working, year(Now) & "-12-31")) + 1
                    If a < 12 And year(rscari!start_working) = year(p_month) Then
                        NilaiHasil = rscari!salary_value * ((DateDiff("m", rscari!start_working, year(Now) & "-12-31")) + 1)
                    Else
                        NilaiHasil = rscari!salary_value * 12
                    End If
                End If
                rscari.Close
            
            ElseIf v_salary_code = "SU-034" Then 'JKK 0.24% or 0.54%
                SQL = "SELECT a.basic_salary, b.company_code, c.jkk " _
                        & "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_jamsostek_detail c ON c.company_code = b.company_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " _
                        & "AND c.jamsostek_code = '" & p_jstk_type & "' AND date(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " _
                        & "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
                        NilaiHasil = (rscari!jkk / 100) * rscari!basic_salary
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
            
            ElseIf v_salary_code = "SU-035" Then 'JKM 0.3%
                SQL = "SELECT a.basic_salary, b.company_code, c.jkm " _
                        & "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_jamsostek_detail c ON c.company_code = b.company_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " _
                        & "AND c.jamsostek_code = '" & p_jstk_type & "' AND date(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " _
                        & "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
                        NilaiHasil = (rscari!jkm / 100) * rscari!basic_salary
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
                
            ElseIf v_salary_code = "SU-036" Then 'Jamsostek Perusahaan 3.7%
                SQL = "SELECT a.basic_salary, b.company_code, c.prs " _
                        & "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_jamsostek_detail c ON c.company_code = b.company_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND b.company_code = '" & p_Company_Code & "' " _
                        & "AND c.jamsostek_code = '" & p_jstk_type & "' AND date(a.salary_date) <= '" & Format(p_month, "yyyy-MM-dd") & "' " _
                        & "ORDER BY a.salary_date DESC LIMIT 1"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    If IIf(IsNull(p_noJamsostek), "", p_noJamsostek) <> "" Then
                        NilaiHasil = (rscari!prs / 100) * rscari!basic_salary
                    Else
                        NilaiHasil = 0
                    End If
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
            
            '+++++++++++++++++++++++++++++++ THR ++++++++++++++++
            ElseIf v_salary_code = "SU-031" Then
                vTgl1 = Format(frm_trans_salary_process.DTPicker_periode_from.Value, "yyyy-MM-dd")
                vTgl2 = Format(frm_trans_salary_process.DTPicker_periode_to.Value, "yyyy-MM-dd")
                
                SQL = "SELECT jmlthr " & _
                        "from t_thr " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND jenis = 0"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = rscari!jmlTHR
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
            '++++++++++++++++++++++++++++++++++++++++
            
            '+++++++++++++++++++++++++++++++ Bonus ++++++++++++++++
            ElseIf v_salary_code = "SU-032" Then
                vTgl1 = Format(frm_trans_salary_process.DTPicker_periode_from.Value, "yyyy-MM-dd")
                vTgl2 = Format(frm_trans_salary_process.DTPicker_periode_to.Value, "yyyy-MM-dd")
                
                SQL = "SELECT SUM(jmlPotong) jmlPotong " & _
                        "from t_income_expense " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND flag_income_expense = 0 AND flag_type = 2"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = IIf(IsNull(rscari!jmlpotong), 0, rscari!jmlpotong)
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
            '++++++++++++++++++++++++++++++++++++++++
            
            '+++++++++++++++++++++++++++++++ Premi ++++++++++++++++
            ElseIf v_salary_code = "SU-029" Then
                vTgl1 = Format(frm_trans_salary_process.DTPicker_periode_from.Value, "yyyy-MM-dd")
                vTgl2 = Format(frm_trans_salary_process.DTPicker_periode_to.Value, "yyyy-MM-dd")
                
                SQL = "SELECT SUM(jmlPotong) jmlPotong " & _
                        "from t_income_expense " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND flag_income_expense = 0 AND flag_type = 0"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = IIf(IsNull(rscari!jmlpotong), 0, rscari!jmlpotong)
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
            '++++++++++++++++++++++++++++++++++++++++
            
            '+++++++++++++++++++++++++++++++ Insentif ++++++++++++++++
            ElseIf v_salary_code = "SU-030" Then
                vTgl1 = Format(frm_trans_salary_process.DTPicker_periode_from.Value, "yyyy-MM-dd")
                vTgl2 = Format(frm_trans_salary_process.DTPicker_periode_to.Value, "yyyy-MM-dd")
                
                SQL = "SELECT SUM(jmlPotong) jmlPotong " & _
                        "from t_income_expense " & _
                        "WHERE employee_code = '" & p_employee_code & "' AND (date(tgltrans) between '" & vTgl1 & "' AND '" & vTgl2 & "') AND flag_income_expense = 0 AND flag_type = 1"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                If rscari.RecordCount > 0 Then
                    NilaiHasil = IIf(IsNull(rscari!jmlpotong), 0, rscari!jmlpotong)
                Else
                    NilaiHasil = 0
                End If
                rscari.Close
            '++++++++++++++++++++++++++++++++++++++++
            
            '++++++++++++++++++++++++++++++++++++ PPh21 ++++++++++++++++++++++++++++++++
            ElseIf v_salary_code = "SU-32" Then 'PPh 5%
                
                SQL = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 1 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-31'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
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
            
            ElseIf v_salary_code = "SU-33" Then 'PPh21 15%
            
                SQL = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 2 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-31'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
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
            
            ElseIf v_salary_code = "SU-34" Then 'PPh21 25%
            
                SQL = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 3 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-31'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
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
                
            ElseIf v_salary_code = "SU-35" Then 'PPh21 30%
            
                SQL = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 4 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-31'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
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
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            '++++++++++++++++++++++++++++++++++++ PPh21 - GROSS UP ++++++++++++++++++++++++++++++++
            ElseIf v_salary_code = "SU-321" Then 'PPh 5%
                
                SQL = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 1 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-311'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
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
            
            ElseIf v_salary_code = "SU-331" Then 'PPh21 15%
            
                SQL = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 2 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-311'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
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
            
            ElseIf v_salary_code = "SU-341" Then 'PPh21 25%
            
                SQL = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 3 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-311'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
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
                
            ElseIf v_salary_code = "SU-351" Then 'PPh21 30%
            
                SQL = "SELECT a.salary_value, d.pph21_under, d.pph21_upper, d.pph21_percentage " _
                        & "FROM h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
                        & "JOIN m_salary_standard c ON a.employee_code = c.employee_code " _
                        & "JOIN (SELECT pph21_code,pph21_number,pph21_under,pph21_upper,pph21_percentage FROM m_pph21_detail " _
                            & "WHERE pph21_number = 4 and pph21_code = '" & p_pph21_type & "' GROUP BY pph21_code) d ON c.pph21_type = d.pph21_code " _
                        & "WHERE a.employee_code = '" & p_employee_code & "' AND LEFT(a.MONTH,7) = '" & Left(p_month, 7) & "' " _
                        & "AND a.salary_code = 'SU-311'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
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
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            ElseIf v_salary_code = "SU-77" Then
                SQL = "SELECT salary_value FROM h_salary " & _
                  "WHERE employee_code = '" & p_employee_code & "' " & _
                  "AND LEFT(MONTH,7) = LEFT('" & p_month & "',7) AND salary_code = 'SU-37'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    NilaiHasil = IIf(IsNull(NilaiHasil), 0, NilaiHasil)
                rscari.Close
                
                SQL = "SELECT s_date1,s_date2,s_date3 INTO vdt1,vdt2,vdt3 FROM s_est WHERE s_number=1"
                CnG.Execute SQL
                
                SQL = "Select (f_count_presence(str_employee_code, vdt1, vdt3)" & _
                    "/(f_count_working_days(str_company_code, str_employee_code, vdt1, vdt2))" & _
                    "* dbl_salary) total"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    NilaiHasil = rscari!total
                rscari.Close
            End If
            
            '+++++++++++++++++++++++++++++++++++++++++++++
            SQL = "UPDATE h_salary SET salary_value = '" & NilaiHasil & "' " & _
                  "WHERE LEFT(MONTH,7) = LEFT('" & p_month & "',7) " & _
                  "AND employee_code = '" & p_employee_code & "' AND salary_code = '" & v_salary_code & "'"
            CnG.Execute SQL
            '+++++++++++++++++++++++++++++++++++++++++++++
                        
            rs_hsalary.MoveNext
        Wend
    End If
End Sub
