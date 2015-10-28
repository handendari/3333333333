VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCheckRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rsa As New clsRSA
Dim h As New HDSN

Public Function masihTrial() As Boolean
Dim rsJam As New ADODB.Recordset
Dim rsReg As New ADODB.Recordset

Dim vtglInstall As String
Dim vtglExpired As String
Dim vtglServer As String
Dim vbtsExpired As String
Dim vUnlimited As Integer

    rsa.m_PrivateKey.Key = pKunci1
    rsa.m_PublicKey.Key = pKunci2
    
    rsa.m_PrivateKey.Length = Len(pKunci1)
    rsa.m_PublicKey.Length = Len(pKunci2)

    '++++++++++++++++++++++++++++++++
    SQL = "SELECT MAX(att_date) FROM h_attendance"
    rsJam.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsJam.RecordCount > 0 Then
        SQL = "SELECT * FROM s_about WHERE id = '" & h.GetSerialNumber & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            tglawal = Format(Now, "yyyy-MM-dd")
        Else
            SQL = "SELECT tgl FROM s_tgl"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                tglawal = Format(rs!tgl, "yyyy-MM-dd")
            Else
                tglawal = Format(Now, "yyyy-MM-dd")
            End If
            rs.Close
        End If
        rscari.Close
        
        vtglServer = Format(IIf(IsNull(rsJam.Fields(0).Value), Now, rsJam.Fields(0).Value), "yyyy-MM-dd")
    End If
    rsJam.Close
    '++++++++==++++++++++++++++++++++++++++
    
    SQL = "SELECT * FROM s_tgl"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    vtglInstall = Format(rsa.Decrypt(rsa.PrivateKey, rs!tgl), "yyyy-MM-dd")
    rs.Close
    
    SQL = "SELECT * FROM s_about WHERE id = '" & h.GetSerialNumber & "'"
    rs.Open SQL, CnG, adOpenForwardOnly
    
    If rs.RecordCount > 0 Then
        SQL = "SELECT tgl2,tgl3,ul FROM s_about WHERE id = '" & h.GetSerialNumber & "'"
        rsReg.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsReg.RecordCount > 0 Then
            If IsNull(rsReg!tgl2) Then
                masihTrial = True
            Else
                If Not IsNull(rsReg!tgl3) Then
                    vtglExpired = rsa.Decrypt(rsa.PrivateKey, IIf(IsNull(rsReg!tgl3), 0, rsReg!tgl3))
                    vUnlimited = rsa.Decrypt(rsa.PrivateKey, IIf(IsNull(rsReg!ul), 0, rsReg!ul))
                        
                    If vUnlimited = 1 Then
                        masihTrial = False
                    Else
                        If IsDate(vtglExpired) Then
                            masihTrial = False
                        Else
                            masihTrial = True
                        End If
                    End If
                Else
                    masihTrial = True
                    vtglInstall = Format(rsa.Decrypt(rsa.PrivateKey, rs!tgl), "yyyy-MM-dd")
                    vbtsExpired = Format(DateAdd("d", 30, vtglInstall), "yyyy-MM-dd")
                    If vtglServer > vbtsExpired Then
                        MsgBox "Your Product Has Expired!" & Chr(13) & _
                                "Please Contact Your Administrator For The Information !", vbCritical, headerMSG
                        vLoadMode = 1
                        frm_etc_register_expired.Show
                    End If
                End If
            End If
        Else
            masihTrial = True
        End If
        rsReg.Close
'        End If
    Else
        masihTrial = True
        vbtsExpired = Format(DateAdd("d", 30, vtglInstall), "yyyy-MM-dd")
        If vtglServer > vbtsExpired Then
            MsgBox "Your Product Has Expired!" & Chr(13) & _
                    "Please Contact Your Administrator For The Information !", vbCritical, headerMSG
            vLoadMode = 1
            frm_etc_register_expired.Show
        End If
    End If
    rs.Close
End Function

Public Function masihBerlaku() As Boolean
Dim rsJam As New ADODB.Recordset
Dim rsReg As New ADODB.Recordset

Dim vtglInstall As String
Dim vtglExpired As String
Dim vtglServer As String
Dim vUnlimited As Integer


    rsa.m_PrivateKey.Key = pKunci1
    rsa.m_PublicKey.Key = pKunci2
    
    rsa.m_PrivateKey.Length = Len(pKunci1)
    rsa.m_PublicKey.Length = Len(pKunci2)

    '++++++++++++++++++++++++++++++++
    SQL = "SELECT MAX(att_date) FROM h_attendance"
    rsJam.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsJam.RecordCount > 0 Then
        SQL = "SELECT * FROM s_about WHERE id = '" & h.GetSerialNumber & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            tglawal = Format(Now, "yyyy-MM-dd")
        Else
            SQL = "SELECT tgl FROM s_tgl"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                tglawal = Format(rs!tgl, "yyyy-MM-dd")
            Else
                tglawal = Format(Now, "yyyy-MM-dd")
            End If
            rs.Close
        End If
        rscari.Close
        
        vtglServer = Format(IIf(IsNull(rsJam.Fields(0).Value), Now, rsJam.Fields(0).Value), "yyyy-MM-dd")
    End If
    rsJam.Close
    '++++++++==++++++++++++++++++++++++++++
    
    SQL = "SELECT tgl1,tgl2,tgl3,ul FROM s_about WHERE id = '" & h.GetSerialNumber & "'"
    rsReg.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsReg.RecordCount > 0 Then
        If IsNull(rsReg!tgl2) Then
            masihBerlaku = False
        Else
            vtglExpired = rsa.Decrypt(rsa.PrivateKey, rsReg!tgl3)
            vUnlimited = rsa.Decrypt(rsa.PrivateKey, rsReg!ul)
            
            If vUnlimited = 1 Then
                masihBerlaku = True
            Else
                If IsDate(vtglExpired) Then
                    If vtglExpired > vtglServer Then
                        masihBerlaku = True
                    Else
                        masihBerlaku = False
                        MsgBox "Your Product Has Expired!" & Chr(13) & _
                                "Please Contact Your Administrator For The Information !", vbCritical, headerMSG
                        vLoadMode = 1
                        frm_etc_register_expired.Show
                    End If
                Else
                    masihBerlaku = False
                    MsgBox "Your Product Has Expired!" & Chr(13) & _
                            "Please Contact Your Administrator For The Information !", vbCritical, headerMSG
                    vLoadMode = 1
                    frm_etc_register_expired.Show
                End If
            End If
        End If
    Else
        masihBerlaku = False
    End If
    rsReg.Close
End Function

Public Function batasKaryawan() As Boolean

    SQL = "SELECT COUNT(*) FROM m_employee WHERE company_code = '" & strCompany & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        If rs.Fields(0).Value >= 3 Then
            batasKaryawan = True
        Else
            batasKaryawan = False
        End If
    Else
        batasKaryawan = False
    End If
    rs.Close
End Function
