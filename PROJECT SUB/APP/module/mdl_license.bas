Attribute VB_Name = "mdl_license"
Public Sub no_license()
MsgBox "This Product has no license!", vbCritical, headerMSG
End
End Sub

Public Function get_license(ByVal str As String) As Boolean
If RC4DeCryptASC(str, pEncryptionPassword) = "µ/”¹;UE" Then
    get_license = True
Else
    get_license = False
End If
End Function

Public Function check_license_old1() As Boolean
Dim h As New HDSN, str_l As String

h.CurrentDrive = Val(DRIVE_ID)
str_l = (h.GetModelNumber & h.GetSerialNumber)
Set h = Nothing

'If HD_SN = RC4DeCryptASC(str_l, pEncryptionPassword) Then
'    check_license = True
'Else
'    check_license = False
'End If
End Function

Public Function check_license() As Boolean
Dim h As New HDSN, str_l As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

rs.Open "select * from m_license order by drive_id asc", CnG, adOpenStatic, adLockReadOnly

While Not rs.EOF

    h.CurrentDrive = Val(rs.Fields("drive_id").Value)
    str_l = (h.GetModelNumber & "" & h.GetSerialNumber)
    If rs.Fields("license_number").Value = RC4DeCryptASC(str_l, pEncryptionPassword) Then
        check_license = True
        Exit Function
    Else
        check_license = False
    End If
    
    rs.MoveNext

Wend
End Function

Public Sub set_license()
Dim rs As New ADODB.Recordset
Dim h As New HDSN

rs.Open "select * from m_drive limit 0,1", CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    h.CurrentDrive = rs.Fields("drive_id").Value
Else
    h.CurrentDrive = DRIVE_ID
End If

Set h = Nothing
CnG.Execute "insert into m_license(drive_id,license_number) values(" & h.CurrentDrive & ",'" _
    & RC4DeCryptASC((h.GetModelNumber & h.GetSerialNumber), pEncryptionPassword) & "')"
End Sub

Public Function RC4DeCryptASC(ByVal plaintxt As String, ByVal psw As String) As String
Dim temp, a, i, j, k
Dim cipherby, cipher As String
Dim arrayEncrypted

i = 0
j = 0
RC4Initialize (psw)

For a = 1 To Len(plaintxt)
    i = (i + 1) Mod 256
    j = (j + sbox(i)) Mod 256
    temp = sbox(i)
    sbox(i) = sbox(j)
    sbox(j) = temp
    
    k = sbox((sbox(i) + sbox(j)) Mod 256)
    
    cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
    cipher = cipher & Chr(cipherby)
Next

RC4DeCryptASC = cipher
End Function

Public Sub RC4Initialize(ByVal strPwd As String)
Dim tempSwap, a, b As Integer
Dim intLength As Integer

intLength = Len(strPwd)
For a = 0 To 255
   Key(a) = Asc(Mid(strPwd, (a Mod intLength) + 1, 1))
   sbox(a) = a
Next

b = 0
For a = 0 To 255
   b = (b + sbox(a) + Key(a)) Mod 256
   tempSwap = sbox(a)
   sbox(a) = sbox(b)
   sbox(b) = tempSwap
Next
End Sub


