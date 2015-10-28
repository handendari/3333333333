VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_str_err As String


Private Function OpenFile(Mode As Integer, str_filename As String) As Boolean
On Error GoTo err

If Mode = 0 Then Open str_filename For Input As #1
If Mode = 1 Then Open str_filename For Output As #1
OpenFile = True

Exit Function
err:
    OpenFile = False
End Function

Private Function CloseFile() As Boolean
On Error GoTo err
Close #1
CloseFile = True

Exit Function
err:
    CloseFile = False
End Function

Public Function readFile(i_FileSys As CFileSys) As Boolean
On Error GoTo err

Dim str_data, str1, str2 As String, i As Integer
Dim intLine As Integer

If Not OpenFile(0, App.Path & "\att\key.txt") Then GoTo err
While Not EOF(1)
    Line Input #1, str_data
    If Not Trim("" & str_data) = "" Then
        i = InStr(1, str_data, "=")
        str1 = LCase(Trim(Left(str_data, i - 1))): str2 = Trim(Right(str_data, Len(str_data) - i))
    End If
    
    'USER_NAME, USER_PASS, HOST_NAME, PORT, DB_NAME, DSN, DESC, DRIVER_PATH, DRIVER_NAME, HD_SN
    If (str1 = "user name") Then
        USER_NAME = str2
    ElseIf (str1 = "user pass") Then
        USER_PASS = str2
    ElseIf (str1 = "host name") Then
        host_name = str2
    ElseIf (str1 = "port") Then
        PORT = str2
    ElseIf (str1 = "db name") Then
        DB_NAME = str2
    ElseIf (str1 = "dsn") Then
        DSN = str2
    ElseIf (str1 = "dsn created") Then
        DSN_CREATED = str2
    ElseIf (str1 = "description") Then
        DESC = str2
    ElseIf (str1 = "driver path") Then
        DRIVER_PATH = str2
    ElseIf (str1 = "driver name") Then
        DRIVER_NAME = str2
    ElseIf (str1 = "license number") Then
        HD_SN = str2
    ElseIf (str1 = "license count") Then
        RUNNING_COUNT = str2
    ElseIf (str1 = "license id") Then
        DRIVE_ID = str2
    ElseIf (str1 = "first run") Then
        FIRST_RUN = Val(str2)
    End If
Wend
CloseFile
readFile = True

Exit Function
err:
    readFile = False
    i_FileSys.str_err = err.Description
End Function

Public Function WriteFile(i_FileSys As CFileSys) As Boolean
On Error GoTo err

'USER_NAME, USER_PASS, HOST_NAME, PORT, DB_NAME, DSN,
'DESC, DRIVER_PATH, DRIVER_NAME, HD_SN
    
If Not OpenFile(1, App.Path & "\att\key.txt") Then GoTo err
Print #1, "user name = " & USER_NAME
Print #1, "user pass = " & USER_PASS
Print #1, "host name = " & host_name
Print #1, "port = " & PORT
Print #1, "db name = " & DB_NAME

Print #1, "dsn = " & DSN
Print #1, "dsn created = " & DSN_CREATED
Print #1, "description = " & DESC
Print #1, "driver path = " & DRIVER_PATH
Print #1, "driver name = " & DRIVER_NAME
Print #1, "license number = " & HD_SN
Print #1, "license count = " & RUNNING_COUNT
Print #1, "license id = " & DRIVE_ID
Print #1, "first run = " & IIf(FIRST_RUN = True, 1, 0)

CloseFile
WriteFile = True

Exit Function
err:
    WriteFile = False
    i_FileSys.str_err = err.Description
End Function

Public Property Get str_err() As String
    str_err = m_str_err
End Property

Public Property Let str_err(i_str_err As String)
    m_str_err = i_str_err
End Property


