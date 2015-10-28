Attribute VB_Name = "mdlPublic"
'Public Const gstrNoDevice = "No Device"
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
