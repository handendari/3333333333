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
On Error GoTo Err

If Mode = 0 Then Open str_filename For Input As #1
If Mode = 1 Then Open str_filename For Output As #1
OpenFile = True

Exit Function
Err:
    OpenFile = False
End Function

Private Function CloseFile() As Boolean
On Error GoTo Err
Close #1
CloseFile = True

Exit Function
Err:
    CloseFile = False
End Function

Public Property Get str_err() As String
    str_err = m_str_err
End Property

Public Property Let str_err(i_str_err As String)
    m_str_err = i_str_err
End Property


