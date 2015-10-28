Attribute VB_Name = "mdl_license"
Public sbox(255) As Variant
Public Key(255) As Variant
Public Const pEncryptionPassword As String = "tomcat070509579"



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

Public Function get_finger_print() As String
'Dim rs As New ADODB.Recordset
Dim h As New HDSN

'rs.Open "select top 1 * from m_drive", Cn, adOpenStatic, adLockBatchOptimistic
'If rs.RecordCount > 0 Then
'    h.CurrentDrive = rs.Fields("drive_id").Value
'Else
'    h.CurrentDrive = DRIVE_ID
'End If

h.CurrentDrive = 0

'Set h = Nothing
'Cn.Execute "insert into m_license(id,drive_id,license_number) values(" _
'    & get_inc_number("m_license", "id", "") & "," & h.CurrentDrive & ",'" _
'    & RC4DeCryptASC((h.GetModelNumber & h.GetSerialNumber), pEncryptionPassword) & "')"


get_finger_print = h.GetModelNumber & h.GetSerialNumber
End Function


