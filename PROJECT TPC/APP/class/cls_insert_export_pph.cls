VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_insert_export_pph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strsql As String
Dim rs2 As New ADODB.Recordset
Dim v_a1 As Double
Dim v_a3 As Double
Dim v_a7 As Double
Dim v_a9 As Double
Dim v_a10 As Double
Dim v_a11 As Double
Dim v_a12 As Double
Dim v_a13 As Double
Dim v_a14 As Double
Dim v_a16 As Double
Dim v_a17 As Double
Dim v_a18 As Double
Dim v_a18_setahun As Double
Dim v_a19 As Double
Dim v_a21 As Double
Dim v_a23 As Double
Dim v_a22 As Double
Dim v_a22b As Double

Dim vJKKJKM As Double
Dim vBonusTHR As Double

Dim vPph5persen As Double
Dim vPph15persen As Double
Dim vPph25persen As Double
Dim vPph30persen As Double
Dim a As Double

Public Sub insert_export_pph(pYear As Integer, pCompany_Code As String, _
                        pEmployee_code As String, nourut As Double)

Dim d1 As String
Dim d2 As String

    d1 = pYear & "-01-01"
    d2 = pYear & "-12-31"
    
    SQL = "CALL spr_pph21('" & d1 & "','" & d2 & "', 1, '" & pCompany_Code & "', 0, 1, '" & pEmployee_code & "', 0,'" & LOGIN_LEVEL & "','" & LOGIN_CODE & "')"
    rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs2.RecordCount > 0 Then
        v_a1 = rs2!basic_salary - rs2!pl_value - rs2!pl_not_value - rs2!alpha_value - _
               rs2!late_value - rs2!other_expense
        v_a3 = rs2!tunj_presence + rs2!tunj_meal + rs2!tunj_transport + rs2!tunj_position + _
               rs2!tunj_shift2 + rs2!tunj_shift3 + rs2!Tunj_lain
        v_a7 = v_a1 + v_a3 + rs2!jkk + rs2!jkm + rs2!uang_makan6
        v_a9 = v_a7 + rs2!thr + rs2!bonus
        
        If (v_a7 * 0.05) > 6000000 Then
            v_a10 = 500000 * 12
        Else
            v_a10 = v_a7 * 0.05
        End If
        
        If v_a10 = 6000000 Then
            v_a11 = 0
        ElseIf v_a10 < 6000000 Then
            If ((0.05 * (rs2!thr + rs2!bonus)) + v_a10) < 6000000 Then
                v_a11 = 0.05 * (rs2!thr + rs2!bonus)
            Else
                v_a11 = 6000000 - v_a10
            End If
        End If
        
        v_a12 = rs2!jht12
        v_a13 = v_a10 + v_a11 + v_a12
        
        v_a14 = v_a9 - v_a13
        
        v_a16 = (v_a14 * 12) / rs2!jml_bulan_kerja
        v_a17 = rs2!ptkp17
        
        If v_a17 > v_a16 Then
            v_a18 = 0
        Else
            v_a18 = v_a16 - v_a17
            v_a18 = Int(v_a18 / 1000) * 1000
        End If
        
'        v_a18_setahun = (v_a18 * rs2!jml_bulan_kerja) / 12
        
        Dim rsPPh As New ADODB.Recordset
        
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 1 AND pph21_code = 'STD'"
        rsPPh.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsPPh.RecordCount > 0 Then
            If v_a18 > rsPPh!pph21_upper Then '50000000
                vPph5persen = (rsPPh!pph21_percentage / 100) * rsPPh!pph21_upper '50000000
            Else
                vPph5persen = (rsPPh!pph21_percentage / 100) * v_a18
            End If
        End If
        rsPPh.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 2 AND pph21_code = 'STD'"
        rsPPh.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsPPh.RecordCount > 0 Then
            If v_a18 <= rsPPh!pph21_under Then '50000000
                vPph15persen = 0
            ElseIf v_a18 < rsPPh!pph21_upper Then '250000000
                vPph15persen = (rsPPh!pph21_percentage / 100) * (v_a18 - rsPPh!pph21_under) '50000000)
            Else
                vPph15persen = (rsPPh!pph21_percentage / 100) * (rsPPh!pph21_upper - rsPPh!pph21_under) '200000000
            End If
        End If
        rsPPh.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 3 AND pph21_code = 'STD'"
        rsPPh.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsPPh.RecordCount > 0 Then
            If v_a18 <= rsPPh!pph21_under Then '500000000
                vPph25persen = 0
            ElseIf v_a18 < rsPPh!pph21_upper Then  '500000000
                vPph25persen = (rsPPh!pph21_percentage / 100) * (v_a18 - rsPPh!pph21_under) '250000000
            Else
                vPph25persen = (rsPPh!pph21_percentage / 100) * (rsPPh!pph21_upper - rsPPh!pph21_under) '250000000
            End If
        End If
        rsPPh.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 4 AND pph21_code = 'STD'"
        rsPPh.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsPPh.RecordCount > 0 Then
            If v_a18 <= rsPPh!pph21_under Then '500000000
                vPph30persen = 0
            Else
                If v_a18 > rsPPh!pph21_under Then '500000000
                    vPph30persen = (rsPPh!pph21_percentage / 100) * (v_a18 - rsPPh!pph21_under) '500000000)
                Else
                    vPph30persen = 0
                End If
            End If
        End If
        rsPPh.Close
        
        v_a19 = (vPph5persen + vPph15persen + vPph25persen + vPph30persen)
        v_a21 = (v_a19 / 12) * rs2!jml_bulan_kerja
                
        v_a22 = (v_a19 / 12) * rs2!jml_bulan_kerja
        v_a22b = (v_a19 / 12) * rs2!jml_bulan_kerja
        
        If v_a21 > v_a22 Then
            v_a23 = v_a21 - v_a22
        Else
            v_a23 = v_a22 - v_a21
        End If
        
        Dim vFlagActive As Integer
        Dim vMarital As String
        
        vJKKJKM = rs2!jkk + rs2!jkm
        vBonusTHR = rs2!thr + rs2!bonus
        vFlagActive = rs2!flag_active
        vMarital = rs2!employee_marital_status
        
        strsql = "INSERT INTO t_pph_export " _
            & "(kode_form,tahun_pajak,pembetulan,nomor_urut,npwp_pegawai,nama_pegawai, " _
            & "alamat_pegawai,company_code,jabatan_pegawai,jenis_kelamin,status_pegawai, " _
            & "status_kawin,flg_asing,status_ptkp,jumlah_tanggungan,masa_perolehan_1, " _
            & "masa_perolehan_2,a1,flg_a2,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14, " _
            & "a15,a16,a17,a18,a19,a20,a21,a22,a22a,a22b,a23,a24,flg_a24,bln_a24) " & _
        "Values " & _
            "('" & nourut & "','" & pYear & "',0,'" & nourut & "'," & _
            "'" & rs2!npwp & "','" & Replace(rs2!EMPLOYEE_NAME, "'", "''") & "','" & rs2!employee_address & "'," & _
            "'" & pCompany_Code & "','" & rs2!title_name & "'," & IIf(rs2!sex = "PEREMPUAN", 2, 1) & "," & _
            "" & IIf(vFlagActive = 0, 2, 1) & "," & IIf(vMarital = "TIDAK KAWIN", 1, 2) & ",0,1," & _
            "'" & rs2!no_of_children & "',1,12," & Round(v_a1, 0) & ",1,0," & Round(v_a3, 0) & "," & _
            "0," & Round(vJKKJKM, 0) & ",0," & Round(v_a7, 0) & "," & Round(vBonusTHR, 0) & "," & _
            "" & Round(v_a9, 0) & "," & Round(v_a10, 0) & "," & Round(v_a11, 0) & "," & Round(v_a12, 0) & "," & _
            "" & Round(v_a13, 0) & "," & Round(v_a14, 0) & ",0," & Round(v_a16, 0) & "," & Round(v_a17, 0) & "," & _
            "" & Round(v_a18, 0) & "," & Round(v_a19, 0) & ",0," & Round(v_a21, 0) & "," & Round(v_a22, 0) & "," & _
            "0," & Round(v_a22b, 0) & "," & Round(v_a23, 0) & ",0,0,'')"
        CnG.Execute strsql
    End If
    rs2.Close
    
End Sub

