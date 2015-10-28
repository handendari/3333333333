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
    MsgBox "Database connection failed..." & Chr(13) & _
            "Please input correct configuration...", vbCritical, "Att & Payroll System"
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
Dim rsJam As New ADODB.Recordset

On Error GoTo ErrHandler
    If DBConnect = True Then
        Koneksi = True
        
        If vLoadMode <> 2 Then
            frm_etc_login.Show 1
        End If
    Else
        Koneksi = False
    End If
    Exit Function
ErrHandler:
    Koneksi = False
End Function

