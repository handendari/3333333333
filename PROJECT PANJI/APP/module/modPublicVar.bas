Attribute VB_Name = "modPublicVar"

Option Explicit

Public Const pEncryptionPassword As String = "ssd12122012"  ' ID Pengenal Software

Public Const pKunci1 As String = "-----BEGIN RSA PRIVATE KEY-----" & vbCrLf & _
                                "ACwZ9KeG9/+tOSEM1T1mcEvScwAi5YvS6hYJ8Pb9A0flttDvPXVnFJ5MalsFFpoE" & vbCrLf & _
                                "NdyyKdlXFzJlQxidtCtpYHn+XxXPeqAwNeSr81ucWF2mbt37rRapgB/eqj3zSWGb" & vbCrLf & _
                                "p0zKmZRRoUqfr9OFero7mp28W/4yNZQsu65/2CqUbWAdhDIAhwDFIbenIDFjUMpv" & vbCrLf & _
                                "XM8/R6QRcCsaJRTtZGd08W/tCXxLxps3F7JWaciPg6Z6xWaar4RBEU1fgMJ8zCqh" & vbCrLf & _
                                "cavhHSyO40qAZlGwUYM+x9+fRcfbzhd3IPDGmwDFgYOHIy+nUEkgaqXPrnHOoMmL" & vbCrLf & _
                                "9d1w/HmSe1FqU2DaPVPfbByScP9AAEAABAIAFzLLEwA8T8ffof9a0d0FItyh3SCz" & vbCrLf & _
                                "z1qOB5oE9Ui5sK54XTpcxiD8ao43bexOhlfh1bq6Hkp/GeH2A+F1KH9mcdie8aq1" & vbCrLf & _
                                "qNzC9Pe2HJWbA9i0Qn0hPDnQyxsk3asQwbg+OWQf/Dc8mVXUYDVS0fjd815vv+qm" & vbCrLf & _
                                "4TuX3IzYljkA6fDgAkYboKMqWnhdpYlCK6cXy1nRU6q/3P4iemF1b639twBTUtSh" & vbCrLf & _
                                "OcIajFoz95za33Mua0jdL+HijQOV7T4zcYEUVFYcW4cR24azlmQu/l4whd0BCvoX" & vbCrLf & _
                                "O9Nmhm0k6X3l4rhH63voWKZp+E/AFJ+nLn0cCc80TSq1Io/pANlLK60m5m7AACQO" & vbCrLf & _
                                "BIKparTPy7XwcIpS40ze4ne1nkboCcZDWBp7JW/KL+uvUBzCDEdbgNepOZeeyji/" & vbCrLf & _
                                "UqMszRVc58FCiPTkpg1fcAQ0La6sGbu/TgM6wzIZaCqNX7/HKO6Ygfd8pO6pDKvL" & vbCrLf & _
                                "aAkbS97kuxmssKHVdUsrj/pxlnS1y0970RPjQh+DCD==" & vbCrLf & _
                                "-----END RSA PRIVATE KEY-----"

Public Const pKunci2 As String = "-----BEGIN CERTIFICATE-----" & vbCrLf & _
                                "DAQAAEAABomJnZAwKydVy+wrEwTmSrEe+SKDaYKQt2mL1CZLNlkp4nw8+eiI+ZRo" & vbCrLf & _
                                "6Q2V1twcXY16qfcKPdzAK7ZYgFvbYFtejqcQUjtKmmOyBX6cVBy8CMwdmM84gsTb" & vbCrLf & _
                                "g6bf9on6iM1+O+Dft6NET+XTBTyL5fNY51eiT2XOVj5lj1IK3vgilTjEBdHUIEsT" & vbCrLf & _
                                "QE8yP6fCWvAmsm0pEHuqGPSzGKhsIrGVYGsEyyQ+5k/M6Lje4x4nMjfQESQdry3u" & vbCrLf & _
                                "MwG29JsFAG8huwR96IoLlRaR6rabjd+YNkmeRZRBZ0ZaoCtNov0LHX77VlnDdXwn" & vbCrLf & _
                                "KTYQUnMZMCS/iaRCifo+p57li0RmtWK=" & vbCrLf & _
                                "-----END CERTIFICATE-----"


Public LOGIN_LEVEL, COMPANY_ACCESS As Integer
Public COMPANY_CODE, LOGIN_NAME, LOGIN_FULLNAME, LOGIN_PASS, LOGIN_CODE, DEPARTMENT_CODE, EMPLOYEE_NAME As String
Public loginTime As Date
Public DATA_LEVEL As Integer

Public sbox(255) As Variant
Public Key(255) As Variant
Public blnUser_Read, blnUser_Add, blnUser_Edit, blnUser_Delete, blnUser_Posting, blnUser_Printing As Boolean

Private mFile As CFileSys

'===
Public Const headerMSG = "SSD TECHNOLOGY"
Public Const DataSize = 459
Public vMachineNumber As Integer
Public vEMachineNumber As Integer

Public BLN_RUNNING As Boolean
Public int_timer_tick As Integer, BLN_AUTO_LOG As Boolean

Public USER_NAME, USER_PASS, host_name, PORT, DB_NAME, DSN, DSN_CREATED As String
Public DESC, DRIVER_PATH, DRIVER_NAME, HD_SN, DRIVE_ID, RUNNING_COUNT As String
Public FG_IP_ADDRESS, FG_PORT_NUMBER As String
Public FG_DEVICE_TYPE As Integer
Public IS_LICENSED, MDI_STATE, FIRST_RUN As Boolean
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim konek               As New clsConnect
'Public Enc              As New clsBlowfish

' ADO connection string
Public CnG              As ADODB.Connection
Public rs               As New ADODB.Recordset
Public rsCari           As New ADODB.Recordset
Public rsList           As New ADODB.Recordset

Public strConn As String, ServerDB As String, UserDB As String
Public nmDB As String, passDB As String, PortDB As String
'Public ComTimb As Integer, BaudRateTimb As Integer, StopBitTimb As Double, ParityTimb As String

Public SQL              As String
Public vSetData         As Integer
Public strJudulProgram  As String
Public vLoadMode        As Integer 'Untuk menentukan expired information
Public vFlagChgPwd      As Integer
Public vFlagUbahPwd     As Integer

Sub Main()
'    With INFO_USER
'        .HostName = GetIPHostName()
'        .ip = GetIPAddress()
'    End With
    getServerVariable
    vLoadMode = 0
    
    If ServerDB = "" Then
        frm_etc_db_config.Show
    Else
        konek.Koneksi
    End If
End Sub

Public Sub getServerVariable()
On Error Resume Next
    If GetInitEntry("CONFIG", "A") <> "" Then
        ServerDB = DecryptINI(GetInitEntry("CONFIG", "A"), pEncryptionPassword)
        UserDB = DecryptINI(GetInitEntry("CONFIG", "B"), pEncryptionPassword)
        passDB = DecryptINI(GetInitEntry("CONFIG", "C"), pEncryptionPassword)
        nmDB = DecryptINI(GetInitEntry("CONFIG", "D"), pEncryptionPassword)
        PortDB = DecryptINI(GetInitEntry("CONFIG", "E"), pEncryptionPassword)
    End If
End Sub

