VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsert_h_salary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strsql As String
Dim rs2 As New ADODB.Recordset
Dim rsCari As New ADODB.Recordset

Dim vJumlahLembur As Double
Dim vTunjLapangan As Double
Dim vIncentiveHadir As Double
Dim vTunjLain As Double
Dim vUangMakan As Double
Dim vPotongAbsen As Double
Dim vGajiKotor As Double
Dim vGajibersih As Double
Dim vPph_DapatKotor As Double
Dim vPph_BiayaJabatan As Double
Dim vTotal_PphPengurangPajak As Double
Dim vTotal_PphDapatKenaPajak As Double
Dim vTotal_PphDapatKenaPajakTahun As Double
Dim vPph5persen As Double
Dim vPph15persen As Double
Dim vPph25persen As Double
Dim vPph30persen As Double
Dim vPph21 As Double
Dim vRound As Double
Dim vGajiBersih_Round As Double
Dim a As Double
Dim vJml_hari_kerja As Double
Dim vJml_Hari_absen As Double

Public Sub Insert_h_salary(pEmployee_code As String, pSex As Integer, pBulan As String, pTgl1 As String, pTgl2 As String, _
                pMarital_status As Integer, pNumber_of_children As Integer, pEnd_working As String, _
                pFlag_active As Integer, pCompany_Code As String)
On Error Resume Next

    vJumlahLembur = 0
    vTunjLapangan = 0
    vIncentiveHadir = 0
    vTunjLain = 0
    vUangMakan = 0
    vPotongAbsen = 0
    vGajiKotor = 0
    vGajibersih = 0
    vPph_DapatKotor = 0
    vPph_BiayaJabatan = 0
    vTotal_PphPengurangPajak = 0
    vTotal_PphDapatKenaPajak = 0
    vTotal_PphDapatKenaPajakTahun = 0
    vPph5persen = 0
    vPph15persen = 0
    vPph25persen = 0
    vPph30persen = 0
    vPph21 = 0
    vRound = 0
    vGajiBersih_Round = 0
    vJml_hari_kerja = 0
    vJml_Hari_absen = 0
    
    'h attandance
    'm salary
    'other income
    'other expense
    'loan
    
' IFNULL((SELECT salary FROM m_salary WHERE employee_code = '" & pEmployee_code & "' ORDER BY salary_date DESC LIMIT 1),0) basic_salary
'    strsql = "SELECT " & _
'        "IFNULL(fn_mc('" & pEmployee_code & "', '" & Format(pBulan, "yyyyMM") & "', '" & IIf(IsNull(pEnd_working), Format(pBulan, "yyyyMM"), pEnd_working) & "', '" & pFlag_active & "'),0) basic_salary," & _
'        "IFNULL((SELECT rate_hour FROM m_salary WHERE employee_code = '" & pEmployee_code & "' ORDER BY salary_date DESC LIMIT 1),0) rate_hour," & _
'        "IFNULL((SELECT allowance_rate FROM m_salary WHERE employee_code = '" & pEmployee_code & "' ORDER BY salary_date DESC LIMIT 1),0) allow_rate," & _
'        "IFNULL((SELECT tunj_lain FROM m_salary WHERE employee_code = '" & pEmployee_code & "' ORDER BY salary_date DESC LIMIT 1),0) tunj_lain," & _
'        "IFNULL((SELECT uangmakan FROM m_salary WHERE employee_code = '" & pEmployee_code & "' ORDER BY salary_date DESC LIMIT 1),0) uangmakan," & _
'        "fn_GetTotalJamLembur('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') total_jam_lembur," & _
'        "fn_jmlHariKerja('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') jmlHariMasuk, " & _
'        "fn_jmlHariAbsent('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') jmlHariAbsent," & _
'        "fn_GetJamsostek_JHT_2persen('" & pEmployee_code & "')* -1 JHT_2_persen," & _
'        "fn_GetJamsostek_jkk_jkm('" & pEmployee_code & "') JKK_JKM_204_persen," & _
'        "fn_GetIncomeothers('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') kurang_bln_lalu," & _
'        "fn_GetTHR('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') jml_thr," & _
'        "fn_GetExpenseothers('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') * - 1 potongan_lain," & _
'        "fn_GetPotonganPinjam('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') * -1 potongan_pinjam," & _
'        "(CASE " & pSex & " WHEN 0 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 1) " & _
'        "Else CASE " & pMarital_status & " WHEN 0 THEN " & _
'            "CASE WHEN " & pNumber_of_children & " = 0 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 1) " & _
'                "WHEN " & pNumber_of_children & " = 1 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 6) " & _
'                "WHEN " & pNumber_of_children & " = 2 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 7) " & _
'                "ELSE (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 8) End " & _
'            "Else CASE WHEN " & pNumber_of_children & " = 0 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 2) " & _
'                "WHEN " & pNumber_of_children & " = 1 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 3) " & _
'                "WHEN " & pNumber_of_children & " = 2 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 4) " & _
'                "ELSE (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 5) End END END)/12 * -1 jml_PTKP"
    
    strsql = "SELECT " & _
        "IFNULL(fn_mc('" & pEmployee_code & "', '" & Format(pBulan, "yyyyMM") & "', '" & IIf(IsNull(pEnd_working), Format(pBulan, "yyyyMM"), pEnd_working) & "', '" & pFlag_active & "', '" & pTgl2 & "'),0) basic_salary," & _
        "IFNULL((SELECT rate_hour FROM m_salary WHERE employee_code = '" & pEmployee_code & "' AND date(entry_date) <= '" & pTgl2 & "' ORDER BY salary_date DESC LIMIT 1),0) rate_hour," & _
        "IFNULL((SELECT allowance_rate FROM m_salary WHERE employee_code = '" & pEmployee_code & "' AND date(entry_date) <= '" & pTgl2 & "' ORDER BY salary_date DESC LIMIT 1),0) allow_rate," & _
        "IFNULL((SELECT tunj_lain FROM m_salary WHERE employee_code = '" & pEmployee_code & "' AND date(entry_date) <= '" & pTgl2 & "' ORDER BY salary_date DESC LIMIT 1),0) tunj_lain," & _
        "IFNULL((SELECT uangmakan FROM m_salary WHERE employee_code = '" & pEmployee_code & "' AND date(entry_date) <= '" & pTgl2 & "' ORDER BY salary_date DESC LIMIT 1),0) uangmakan," & _
        "fn_GetTotalJamLembur('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') total_jam_lembur," & _
        "fn_jmlHariKerja('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') jmlHariMasuk, " & _
        "fn_jmlHariAbsent('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') jmlHariAbsent," & _
        "fn_GetJamsostek_JHT_2persen('" & pEmployee_code & "','" & pTgl2 & "')* -1 JHT_2_persen," & _
        "fn_GetJamsostek_jkk_jkm('" & pEmployee_code & "','" & pTgl2 & "') JKK_JKM_204_persen," & _
        "fn_GetJP('" & pEmployee_code & "','" & pTgl2 & "')* -1 JP," & _
        "fn_GetBPJSKesehatan('" & pEmployee_code & "','" & pTgl2 & "')* -1 BPJS_kesehatan," & _
        "fn_GetIncomeothers('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') kurang_bln_lalu," & _
        "fn_GetTHR('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') jml_thr," & _
        "fn_GetExpenseothers('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') * - 1 potongan_lain," & _
        "fn_GetPotonganPinjam('" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "') * -1 potongan_pinjam,"
    
    strsql = strsql & _
        "(CASE " & pSex & " WHEN 0 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 1) " & _
        "Else CASE " & pMarital_status & " WHEN 0 THEN " & _
            "CASE WHEN " & pNumber_of_children & " = 0 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 1) " & _
                "WHEN " & pNumber_of_children & " = 1 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 6) " & _
                "WHEN " & pNumber_of_children & " = 2 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 7) " & _
                "ELSE (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 8) End " & _
            "Else CASE WHEN " & pNumber_of_children & " = 0 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 2) " & _
                "WHEN " & pNumber_of_children & " = 1 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 3) " & _
                "WHEN " & pNumber_of_children & " = 2 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 4) " & _
                "ELSE (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 5) End END END)/12 * -1 jml_PTKP"
                
'        "(CASE " & pSex & " WHEN 0 THEN (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 1) " & _
'        "Else CASE " & pMarital_status & " WHEN 0 THEN " &
'                "(SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 1) " &
'            "Else CASE WHEN " & pNumber_of_children & " = 0 THEN " &
'                    "(SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 2) " &
'                "WHEN " & pNumber_of_children & " = 1 THEN " &
'                    "(SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 3) " &
'                "WHEN " & pNumber_of_children & " = 2 THEN " &
'                    "(SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 4) " &
'                "ELSE (SELECT ptkp_value FROM m_ptkp WHERE ptkp_number = 5) " &
'                "End End END)/12 * -1 jml_PTKP"
    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs2.RecordCount > 0 Then
        Dim rslembur As New ADODB.Recordset
        Dim rsbasic As New ADODB.Recordset
        Dim a As Integer
        Dim c As Integer
        Dim b As String
        
        vJml_hari_kerja = rs2!jmlHariMasuk
        vJml_Hari_absen = rs2!jmlHariAbsent
        
        vUangMakan = vJml_hari_kerja * rs2!uangmakan
        
        vTunjLapangan = rs2!allow_rate * vJml_hari_kerja
        vTunjLain = IIf(IsNull(rs2!tunj_lain), 0, rs2!tunj_lain)
        
        strsql = "SELECT a.employee_code,a.managerial_level,b.flag_incentive " _
                & "FROM m_employee a LEFT JOIN m_salary b ON a.employee_code = b.employee_code " _
                & "WHERE a.employee_code = '" & pEmployee_code & "' " _
                & "ORDER BY salary_date DESC LIMIT 1"
        rslembur.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rslembur.RecordCount > 0 Then
            rslembur.MoveFirst
            While Not rslembur.EOF
                a = IIf(IsNull(rslembur!managerial_level), 0, rslembur!managerial_level)
                c = IIf(IsNull(rslembur!flag_incentive), 0, rslembur!flag_incentive)
   ' c=0 incentiv =0
   'c=1 dpt incentiv
   'a=0,incetiv + lembut
   'a=1, tdk dpt sama sekali
   
                If c = 0 And a = 0 Then
                    v_Incentivehadir = 0
                    vJumlahLembur = rs2!rate_hour * rs2!total_jam_lembur
                ElseIf c = 1 And a = 1 Then
                    vJumlahLembur = 0
                    v_Incentivehadir = 0
                ElseIf c = 0 And a = 1 Then
                    vJumlahLembur = 0
                    v_Incentivehadir = 0
                ElseIf c = 1 And a = 0 Then
                    vJumlahLembur = rs2!rate_hour * rs2!total_jam_lembur
                    'vJumlahLembur = 0
                    If vJml_hari_kerja < 26 Then
                        vIncentiveHadir = 0
                    ElseIf vJml_hari_kerja >= 26 And vJml_hari_kerja < 28 Then
                        vIncentiveHadir = ((rs2!basic_salary + vTunjLapangan + vJumlahLembur) / 28) * 1
                    Else
                        vIncentiveHadir = ((rs2!basic_salary + vTunjLapangan + vJumlahLembur) / 28) * 2
                    End If
                    'v_Incentivehadir = 0
'                Else
'                    vJumlahLembur = rs2!rate_hour * rs2!total_jam_lembur
'
'                    If vJml_hari_kerja < 26 Then
'                        vIncentiveHadir = 0
'                    ElseIf vJml_hari_kerja >= 26 And vJml_hari_kerja < 28 Then
'                        vIncentiveHadir = ((rs2!basic_salary + vTunjLapangan + vJumlahLembur) / 28) * 1
'                    Else
'                        vIncentiveHadir = ((rs2!basic_salary + vTunjLapangan + vJumlahLembur) / 28) * 2
'                    End If
                End If
            rslembur.MoveNext
            Wend
        End If
        rslembur.Close

'        If rs2!jmlHariMasuk < 26 Then
'            vIncentiveHadir = 0
'        ElseIf rs2!jmlHariMasuk > 26 And rs2!jmlHariMasuk < 28 Then
'            vIncentiveHadir = ((rs2!basic_salary + vTunjLapangan + vJumlahLembur) / 28) * 1
'        Else
'            vIncentiveHadir = ((rs2!basic_salary + vTunjLapangan + vJumlahLembur) / 28) * 2
'        End If
        
        vPotongAbsen = vJml_Hari_absen / 25 * rs2!basic_salary * -1
        
        vGajiKotor = (rs2!basic_salary + vJumlahLembur + vTunjLapangan + vIncentiveHadir + rs2!kurang_bln_lalu + vTunjLain + rs2!jml_thr + vUangMakan) _
            + (vPotongAbsen + rs2!potongan_lain)
                
        vPph_DapatKotor = vGajiKotor + rs2!JKK_JKM_204_persen
        
        If (vPph_DapatKotor * 12 * 0.05) > 6000000 Then
            vPph_BiayaJabatan = 500000 * -1
        Else
            vPph_BiayaJabatan = vPph_DapatKotor * 0.05 * -1
        End If
        
        vTotal_PphPengurangPajak = vPph_BiayaJabatan + rs2!JHT_2_persen + rs2!jml_ptkp + rs2!JP
        vTotal_PphDapatKenaPajak = vPph_DapatKotor + vTotal_PphPengurangPajak
        vTotal_PphDapatKenaPajakTahun = (vTotal_PphDapatKenaPajak * 12)
        
        Dim rspph As New ADODB.Recordset
        
        '+++++++++++++++++++++++++++ Perhitungan PPh21 Tanpa THR +++++++++++++++++++++++++++++++++++++++
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 1"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If vTotal_PphDapatKenaPajakTahun < 0 Then
                vPph5persen = 0
            Else
                If vTotal_PphDapatKenaPajakTahun > rspph!pph21_upper Then '50000000
                    vPph5persen = (rspph!pph21_percentage / 100) * rspph!pph21_upper '50000000
                Else
                    vPph5persen = (rspph!pph21_percentage / 100) * vTotal_PphDapatKenaPajakTahun
                End If
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 2"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If vTotal_PphDapatKenaPajakTahun < 0 Then
                vPph15persen = 0
            Else
                If vTotal_PphDapatKenaPajakTahun <= rspph!pph21_under Then '50000000
                    vPph15persen = 0
                ElseIf vTotal_PphDapatKenaPajakTahun < rspph!pph21_upper Then '250000000
                    vPph15persen = (rspph!pph21_percentage / 100) * (vTotal_PphDapatKenaPajakTahun - rspph!pph21_under) '50000000)
                Else
                    vPph15persen = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '200000000
                End If
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 3"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If vTotal_PphDapatKenaPajakTahun < 0 Then
                vPph25persen = 0
            Else
                If vTotal_PphDapatKenaPajakTahun <= rspph!pph21_under Then '500000000
                    vPph25persen = 0
                ElseIf vTotal_PphDapatKenaPajakTahun < rspph!pph21_upper Then  '500000000
                    vPph25persen = (rspph!pph21_percentage / 100) * (vTotal_PphDapatKenaPajakTahun - rspph!pph21_under) '250000000
                Else
                    vPph25persen = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '250000000
                End If
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 4"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If vTotal_PphDapatKenaPajakTahun < 0 Then
                vPph35persen = 0
            Else
                If vTotal_PphDapatKenaPajakTahun <= rspph!pph21_under Then '500000000
                    vPph30persen = 0
                Else
                    If vTotal_PphDapatKenaPajakTahun > rspph!pph21_under Then '500000000
                        vPph30persen = (rspph!pph21_percentage / 100) * (vTotal_PphDapatKenaPajakTahun - rspph!pph21_under) '500000000)
                    Else
                        vPph30persen = 0
                    End If
                End If
            End If
        End If
        rspph.Close
        '+++++++++++++++++++++++++++ END Perhitungan PPh21 Tanpa THR ++++++++++++++++++++++++++++++++++++
        
'        '+++++++++++++++++++++++++++++ Perhitungan PPh21 Dengan THR +++++++++++++++++++++++++++++++++++++
'        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 1"
'        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'        If rspph.RecordCount > 0 Then
'            If vTotal_PphDapatKenaPajakTahun < 0 Then
'                vPph5persen = 0
'            Else
'                If vTotal_PphDapatKenaPajakTahun > rspph!pph21_upper Then '50000000
'                    vPph5persen = (rspph!pph21_percentage / 100) * rspph!pph21_upper '50000000
'                Else
'                    vPph5persen = (rspph!pph21_percentage / 100) * vTotal_PphDapatKenaPajakTahun
'                End If
'            End If
'        End If
'        rspph.Close
'
'        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 2"
'        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'        If rspph.RecordCount > 0 Then
'            If vTotal_PphDapatKenaPajakTahun < 0 Then
'                vPph15persen = 0
'            Else
'                If vTotal_PphDapatKenaPajakTahun <= rspph!pph21_under Then '50000000
'                    vPph15persen = 0
'                ElseIf vTotal_PphDapatKenaPajakTahun < rspph!pph21_upper Then '250000000
'                    vPph15persen = (rspph!pph21_percentage / 100) * (vTotal_PphDapatKenaPajakTahun - rspph!pph21_under) '50000000)
'                Else
'                    vPph15persen = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '200000000
'                End If
'            End If
'        End If
'        rspph.Close
'
'        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 3"
'        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'        If rspph.RecordCount > 0 Then
'            If vTotal_PphDapatKenaPajakTahun < 0 Then
'                vPph25persen = 0
'            Else
'                If vTotal_PphDapatKenaPajakTahun <= rspph!pph21_under Then '500000000
'                    vPph25persen = 0
'                ElseIf vTotal_PphDapatKenaPajakTahun < rspph!pph21_upper Then  '500000000
'                    vPph25persen = (rspph!pph21_percentage / 100) * (vTotal_PphDapatKenaPajakTahun - rspph!pph21_under) '250000000
'                Else
'                    vPph25persen = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '250000000
'                End If
'            End If
'        End If
'        rspph.Close
'
'        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 4"
'        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'        If rspph.RecordCount > 0 Then
'            If vTotal_PphDapatKenaPajakTahun < 0 Then
'                vPph35persen = 0
'            Else
'                If vTotal_PphDapatKenaPajakTahun <= rspph!pph21_under Then '500000000
'                    vPph30persen = 0
'                Else
'                    If vTotal_PphDapatKenaPajakTahun > rspph!pph21_under Then '500000000
'                        vPph30persen = (rspph!pph21_percentage / 100) * (vTotal_PphDapatKenaPajakTahun - rspph!pph21_under) '500000000)
'                    Else
'                        vPph30persen = 0
'                    End If
'                End If
'            End If
'        End If
'        rspph.Close
'        '+++++++++++++++++++++++++++ END Perhitungan PPh21 Tanpa THR ++++++++++++++++++++++++++++++++++++
        
        vPph21 = (vPph5persen + vPph15persen + vPph25persen + vPph30persen) / 12 * -1
        
        vGajibersih = Round(vGajiKotor) + rs2!JHT_2_persen + rs2!potongan_pinjam + Round(vPph21) + rs2!JP + rs2!BPJS_kesehatan
        
        a = vGajibersih Mod 1000
        vRound = a * -1
'        vRound = 1000 - a
        
        vGajiBersih_Round = vGajibersih + vRound
        
'        If pEmployee_code = "08041000062" Then
'            MsgBox "iki 08041000062" & "-" & vJml_Hari_absen & vTunjLapangan
'        End If
        
        strsql = "DELETE FROM h_salary_new WHERE month = '" & pBulan & "' " & _
                    "AND employee_code = '" & pEmployee_code & "'"
        CnG.Execute strsql
        
        strsql = "INSERT INTO h_salary_new " & _
            "(month,employee_code,date_from,date_to," & _
            "salary_basic,rate_hour,allowance_rate,uang_makan,jml_jam_lembur," & _
            "jml_lembur,jml_hari_absen,description,tunj_lap,insentive_hadir,tunj_lain," & _
            "krg_bln_lalu,thr,pot_absensi,pot_others,gaji_kotor,pot_pinjaman," & _
            "pot_jamsostek,pjk_jkk_jkm_204,tot_dpt_kotor,pjk_tunj_jabatan,pjk_ptkp,tot_pengurang_pjk," & _
            "tot_dpt_kena_pajak,tot_dpt_kena_pajak_tahun,pph_5_persen,pph_15_persen,pph_25_persen," & _
            "pph_30_persen,pph21,round,gaji_bersih,company_code,pot_jp,pot_bpjs_kes) " & _
        "Values " & _
            "('" & pBulan & "','" & pEmployee_code & "','" & pTgl1 & "','" & pTgl2 & "'," & _
            "'" & rs2!basic_salary & "','" & rs2!rate_hour & "','" & rs2!allow_rate & "','" & vUangMakan & "','" & rs2!total_jam_lembur & "'," & _
            "'" & vJumlahLembur & "','" & vJml_Hari_absen & "',''," & vTunjLapangan & ",'" & vIncentiveHadir & "','" & vTunjLain & "'," & _
            "'" & rs2!kurang_bln_lalu & "','" & rs2!jml_thr & "','" & vPotongAbsen & "','" & rs2!potongan_lain & "','" & vGajiKotor & "','" & rs2!potongan_pinjam & "'," & _
            "'" & rs2!JHT_2_persen & "', '" & rs2!JKK_JKM_204_persen & "','" & vPph_DapatKotor & "','" & vPph_BiayaJabatan & "'," & _
            "'" & rs2!jml_ptkp & "', '" & vTotal_PphPengurangPajak & "', '" & vTotal_PphDapatKenaPajak & "'," & _
            "'" & vTotal_PphDapatKenaPajakTahun & "', '" & vPph5persen & "','" & vPph15persen & "'," & _
            "'" & vPph25persen & "','" & vPph30persen & "','" & vPph21 & "','" & vRound & "','" & vGajiBersih_Round & "','" & pCompany_Code & "'," & _
            "'" & rs2!JP & "','" & rs2!BPJS_kesehatan & "')"
        CnG.Execute strsql
        
'        If pEmployee_code = "08041000062" Then
'            MsgBox "iki 08041000062" & "-" & vJml_Hari_absen & vTunjLapangan
'        End If
    End If
    rs2.Close
    
End Sub


