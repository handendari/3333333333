Attribute VB_Name = "Module1"
Public Const headerMSG = "FGP"
Public strConn As String
Public Cn As New ADODB.Connection
Public lng_start, lng_end As Long
Public Const DATASIZE = 459
Public vMachineNumber As Integer
Public gGetState As Boolean


Function ErrorPrint(aErrorCode As Long) As String
Select Case aErrorCode
    Case 1
        ErrorPrint = "SUCCESS"
    Case 4
        ErrorPrint = "ERR_INVALID_PARAM"
    Case 0
        ErrorPrint = "ERR_NO_DATA"
    Case -4
        ErrorPrint = "ERROR_NO_SPACE"
    Case -3
        ErrorPrint = "ERROR_SIZE "
    Case -2
        ErrorPrint = "ERROR_IO"
    Case -1
        ErrorPrint = "ERROR_NOT_INIT"
    Case -100
        ErrorPrint = "ERROR_UNSUPPORT"
End Select
End Function

Sub main()
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
& App.Path & "\datEnrollDat.mdb;Persist Security Info=False"

Cn.ConnectionString = strConn
Cn.Open

Call get_en_number_setting
mdi_fgp.Show
End Sub

Public Sub get_en_number_setting()
Dim rs As New ADODB.Recordset
rs.Open "select * from s_enroll_number where no_urut=1", Cn, adOpenKeyset, adLockOptimistic

If rs.RecordCount > 0 Then
    lng_start = rs.Fields("en_number_start").Value
    lng_end = rs.Fields("en_number_end").Value
Else
    lng_start = 1
    lng_end = 100
End If
End Sub
