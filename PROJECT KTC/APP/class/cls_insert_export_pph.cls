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
Dim vPph5persen As Double
Dim vPph15persen As Double
Dim vPph25persen As Double
Dim vPph30persen As Double
Dim a As Double

Public Sub insert_export_pph(pYear As Integer, pCompany_Code As String, _
                        pEmployee_code As String, nourut As Double)
       
    strsql = "SELECT spr_pph21_export('" & pEmployee_code & "'," & pYear & ",'" & pEmployee_code & "')"
    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs2.RecordCount > 0 Then
        v_a1 = rs2!salary_basic
        v_a3 = (rs2!jml_lembur + rs2!tunj_lap + rs2!insentive_hadir) + (rs2!pot_absensi + rs2!pot_others)
        v_a7 = v_a1 + v_a3 + rs2!pjk_jkk_jkm_204
        v_a9 = v_a7 + rs2!thr
        
        If (v_a9 * 0.05) > 6000000 Then
            v_a10 = roundDown(500000 * 12)
        Else
            v_a10 = roundDown(v_a9 * 0.05)
        End If
        
        If v_a10 = 6000000 Then
            v_a11 = 0
        ElseIf v_a10 < 6000000 Then
            If ((0.05 * rs2!thr) + v_a10) < 6000000 Then
                v_a11 = roundDown(0.05 * rs2!thr)
            Else
                v_a11 = roundDown(6000000 - v_a10)
            End If
        End If
        
        v_a12 = rs2!pot_jamsostek * -1
        v_a13 = v_a10 + v_a11 + v_a12
        
        v_a14 = v_a9 - v_a13
        
        v_a16 = v_a14
        v_a17 = (rs2!pjk_ptkp * -1) * 12
        
        If v_a17 > v_a16 Then
            v_a18 = 0
        Else
            v_a18 = Round(v_a16 - v_a17, 0)
            v_a18 = v_a18 - (v_a18 Mod 1000)
        End If
        
        v_a18_setahun = v_a18
'        v_a18_setahun = (v_a18 * rs2!jml_bln_kerja) / 12
        
        Dim rspph As New ADODB.Recordset
        
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 1"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If v_a18_setahun > rspph!pph21_upper Then '50000000
                vPph5persen = (rspph!pph21_percentage / 100) * rspph!pph21_upper '50000000
            Else
                vPph5persen = (rspph!pph21_percentage / 100) * v_a18_setahun
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 2"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If v_a18_setahun <= rspph!pph21_under Then '50000000
                vPph15persen = 0
            ElseIf v_a18_setahun < rspph!pph21_upper Then '250000000
                vPph15persen = (rspph!pph21_percentage / 100) * (v_a18_setahun - rspph!pph21_under) '50000000)
            Else
                vPph15persen = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '200000000
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 3"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If v_a18_setahun <= rspph!pph21_under Then '500000000
                vPph25persen = 0
            ElseIf v_a18_setahun < rspph!pph21_upper Then  '500000000
                vPph25persen = (rspph!pph21_percentage / 100) * (v_a18_setahun - rspph!pph21_under) '250000000
            Else
                vPph25persen = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '250000000
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21 WHERE pph21_number = 4"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If v_a18_setahun <= rspph!pph21_under Then '500000000
                vPph30persen = 0
            Else
                If v_a18_setahun > rspph!pph21_under Then '500000000
                    vPph30persen = (rspph!pph21_percentage / 100) * (v_a18_setahun - rspph!pph21_under) '500000000)
                Else
                    vPph30persen = 0
                End If
            End If
        End If
        rspph.Close
        
        v_a19 = (vPph5persen + vPph15persen + vPph25persen + vPph30persen)
        v_a21 = v_a19
        
'        If v_a17 > v_a16 Then
'            v_a22 = 0
'        Else
'            v_a22 = rs2!pph21 * -1
'        End If
        v_a22 = v_a19
            
        If v_a21 > v_a22 Then
            v_a23 = v_a21 - v_a22
        Else
            v_a23 = v_a22 - v_a21
        End If
        
        strsql = "INSERT INTO t_pph_export " _
            & "(kode_form,tahun_pajak,pembetulan,nomor_urut,npwp_pegawai,nama_pegawai, " _
            & "alamat_pegawai,company_code,jabatan_pegawai,jenis_kelamin,status_pegawai, " _
            & "status_kawin,flg_asing,status_ptkp,jumlah_tanggungan,masa_perolehan_1, " _
            & "masa_perolehan_2,a1,flg_a2,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14, " _
            & "a15,a16,a17,a18,a19,a20,a21,a22,a22a,a22b,a23,a24,flg_a24,bln_a24) " & _
        "Values " & _
            "(" & nourut & ",'" & pYear & "',0,'" & nourut & "'," & _
            "'" & rs2!npwp & "','" & Replace(rs2!employee_name, "'", "''") & "','" & rs2!Address & "'," & _
            "'" & pCompany_Code & "','" & rs2!title_name & "'," & rs2!jenis_kelamin & "," & _
            "" & rs2!status_pegawai & "," & rs2!marital_status & ",0,1," & _
            "" & rs2!number_of_children & ",1,12," & Round(v_a1, 0) & ",1,0," & Round(v_a3, 0) & "," & _
            "0," & Round(rs2!pjk_jkk_jkm_204, 0) & ",0," & Round(v_a7, 0) & "," & Round(rs2!thr, 0) & "," & _
            "" & Round(v_a9, 0) & "," & Round(v_a10, 0) & "," & Round(v_a11, 0) & "," & Round(v_a12, 0) & "," & _
            "" & Round(v_a13, 0) & "," & Round(v_a14, 0) & ",0," & Round(v_a16, 0) & "," & Round(v_a17, 0) & "," & _
            "" & Round(v_a18, 0) & "," & Round(v_a19, 0) & ",0," & Round(v_a21, 0) & "," & Round(v_a22, 0) & "," & _
            "0," & Round(v_a22, 0) & "," & Round(v_a23, 0) & ",0,0,'')"
        CnG.Execute strsql
    End If
    rs2.Close
    
End Sub




