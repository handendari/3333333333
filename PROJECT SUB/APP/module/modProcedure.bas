Attribute VB_Name = "modProcedure"
Option Explicit
Private sDefInitFileName As String
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetInitEntry(ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "", Optional ByVal sInitFileName As String = "") As String

'This Function Reads In a String From The Init File.
'Returns Value From Init File or sDefault If No Value Exists.
'sDefault Defaults to an Empty String ("").
'Creates and Uses sDefInitFileName (AppPath\AppEXEName.Ini)
'if sInitFileName Parameter Is Not Passed In.

Dim sBuffer As String
Dim sInitFile As String

    'If Init Filename NOT Passed In
    If Len(sInitFileName) = 0 Then
        'If Static Init FileName NOT Already Created
        If Len(sDefInitFileName) = 0 Then
            'Create Static Init FileName
            sDefInitFileName = App.Path
            If Right$(sDefInitFileName, 1) <> "\config\" Then
                sDefInitFileName = sDefInitFileName & "\config\"
            End If
            sDefInitFileName = sDefInitFileName & "config.ssd"
        End If
        sInitFile = sDefInitFileName
    Else    'If Init Filename Passed In
        sInitFile = sInitFileName
    End If
    
    sBuffer = String$(2048, " ")
    GetInitEntry = Left$(sBuffer, GetPrivateProfileString(sSection, ByVal sKeyName, sDefault, sBuffer, Len(sBuffer), sInitFile))

End Function

Public Function SetInitEntry(ByVal sSection As String, Optional ByVal sKeyName As String, Optional ByVal sValue As String, Optional ByVal sInitFileName As String = "") As Long
Dim sInitFile As String

    'If Init Filename NOT Passed In
    If Len(sInitFileName) = 0 Then
        'If Static Init FileName NOT Already Created
        If Len(sDefInitFileName) = 0 Then
            'Create Static Init FileName
            sDefInitFileName = App.Path
            If Right$(sDefInitFileName, 1) <> "\config\" Then
                sDefInitFileName = sDefInitFileName & "\config\"
            End If
            sDefInitFileName = sDefInitFileName & "config.ssd"
        End If
        sInitFile = sDefInitFileName
    Else    'If Init Filename Passed In
        sInitFile = sInitFileName
    End If
    
    If Len(sKeyName) > 0 And Len(sValue) > 0 Then
        SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, ByVal sValue, sInitFile)
    ElseIf Len(sKeyName) > 0 Then
        SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, vbNullString, sInitFile)
    Else
        SetInitEntry = WritePrivateProfileString(sSection, vbNullString, vbNullString, sInitFile)
    End If

End Function

'Procedure used to highlight text when focus
Public Sub HLText(ByRef sText)
    On Error Resume Next
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
End Sub

'Procedure used to clear the text content
Public Sub clearText(ByRef sForm As Form)
    Dim CONTROL As CONTROL
    For Each CONTROL In sForm.Controls
        If (TypeOf CONTROL Is TextBox) Then CONTROL = vbNullString
    Next CONTROL
    Set CONTROL = Nothing
End Sub

'Procedure used to clear the text content
Public Sub LockInput(ByRef sForm As Form, ByVal bolLock As Boolean, Optional bolTabStop As Boolean)
    On Error Resume Next
    Dim CONTROL As CONTROL
    For Each CONTROL In sForm.Controls
       CONTROL.Locked = bolLock
    Next CONTROL
    Set CONTROL = Nothing
End Sub

'Procedure used to center form
Public Sub centerForm(ByRef sForm As Form, ByVal sHeight As Integer, ByVal sWidth As Integer)
    sForm.Move (sWidth - sForm.Width) / 2, (sHeight - sForm.Height) / 2
End Sub
'Procedure used to center object horizontal
Public Sub center_obj_horizontal(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.Left = (sParentObj - sMoveObj.Width) / 2
End Sub
'Procedure used to center vertical
Public Sub center_obj_vertical(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.Top = (sParentObj.Height - sMoveObj.Height) / 2
End Sub



