VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim rsa As New clsRSA

Friend Function DBConnect() As Boolean
    On Error GoTo ErrHandler
    getServerVariable
    strConn = "DATABASE=" & nmDB & ";SERVER=" & ServerDB & ";port=" & PortDB & ";UID=" & UserDB & ";PWD=" & passDB & ";provider=MSDASQL.1;DRIVER={MySQL ODBC 5.1 Driver};"
    Set CnG = New ADODB.Connection
    With CnG
       .CommandTimeout = 100
       .ConnectionString = strConn
       .Open
       .CursorLocation = adUseClient
    End With
    DBConnect = True
    Exit Function
    
ErrHandler:
    MsgBox "Koneksi Database Gagal..." & Chr(13) & _
            "Pastikan Konfigurasi Database Benar...", vbCritical, "Att & Payroll System"
    frm_etc_db_config.Show
    DBConnect = False
End Function

Friend Function DBDisConnect() As Boolean
    On Error GoTo ErrHandler
    CnG.Close
    Set CnG = Nothing
    DBDisConnect = True
    Exit Function
ErrHandler:
    DBDisConnect = False
End Function

Public Function Koneksi() As Boolean
Dim h As HDSN
Dim clsReg As New clsCheckRegister
Dim rsJam As New ADODB.Recordset
Dim vDatabaseType As String

On Error GoTo ErrHandler
    If DBConnect = True Then
        Koneksi = True
        
        '++++++++++++++++++++++++++++++++++++++++
        rsa.m_PrivateKey.Key = pKunci1
        rsa.m_PublicKey.Key = pKunci2
        
        rsa.m_PrivateKey.Length = Len(pKunci1)
        rsa.m_PublicKey.Length = Len(pKunci2)
        '+========================================
    
        SQL = "SELECT * FROM s_tgl"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount = 0 Then
            SQL = "SELECT NOW() tglServer"
            rsJam.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            SQL = "INSERT INTO s_tgl(tgl) " & _
                  "VALUE ('" & rsa.Encrypt(rsa.PublicKey, rsJam!tglServer) & "')"
            CnG.Execute SQL
            rsJam.Close
        End If
        rs.Close
        
        vDatabaseType = IIf(TypeDB = 0, "SERVER DATABASE", "LOCAL DATABASE")
        If clsReg.masihTrial = True Then
            If vLoadMode = 1 Then
                strJudulProgram = "ATTENDANCE & PAYROLL SYSTEM 3.0 (TRIAL VERSION) - " & vDatabaseType
            ElseIf vLoadMode = 2 Then
                strJudulProgram = "ATTENDANCE & PAYROLL SYSTEM 3.0 (TRIAL VERSION) - " & vDatabaseType
            Else
                strJudulProgram = "ATTENDANCE & PAYROLL SYSTEM 3.0 (TRIAL VERSION) - " & vDatabaseType
                
'                frm_etc_login.timer1.Enabled = True
                frm_etc_login.Show 1
            End If
        Else
            If clsReg.masihBerlaku = False Then
                If vLoadMode = 1 Then
                    strJudulProgram = "ATTENDANCE & PAYROLL SYSTEM 3.0 (TRIAL VERSION) - " & vDatabaseType
                ElseIf vLoadMode = 2 Then
                    strJudulProgram = "ATTENDANCE & PAYROLL SYSTEM 3.0 (TRIAL VERSION) - " & vDatabaseType
                Else
                    strJudulProgram = "ATTENDANCE & PAYROLL SYSTEM 3.0 (TRIAL VERSION) - " & vDatabaseType
                                       
'                    frm_etc_login.timer1.Enabled = True
                    frm_etc_login.Show 1
                End If
            Else
                If vLoadMode = 1 Then
                    strJudulProgram = "ATTENDANCE & PAYROLL SYSTEM 3.0 - " & vDatabaseType
                ElseIf vLoadMode = 2 Then
                    strJudulProgram = "ATTENDANCE & PAYROLL SYSTEM 3.0 - " & vDatabaseType
                Else
                    strJudulProgram = "ATTENDANCE & PAYROLL SYSTEM 3.0 - " & vDatabaseType
                    
'                    frm_etc_login.timer1.Enabled = True
                    frm_etc_login.Show 1
                End If
            End If
        End If
    Else
        Koneksi = False
    End If
    Exit Function
ErrHandler:
    Koneksi = False
End Function

