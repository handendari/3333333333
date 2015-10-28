Attribute VB_Name = "Error"
Option Explicit
Public ErrorLocation As String

Public Sub ErrorMsg(Error As String, Source As String)
Dim fso As New FileSystemObject, j As Integer, i As Integer
Dim fsotext As TextStream, FileName As String, Directory As String
Directory = Environ("SystemRoot")
FileName = "ERROR.666"
Set fsotext = fso.OpenTextFile(Directory + "\" + FileName, ForAppending, True, TristateUseDefault)
j = fsotext.Line
If j >= 666 Then
    fsotext.Close
    Dim StrBuffer(700) As String
    Set fsotext = fso.OpenTextFile(Directory + "\" + FileName, ForReading, True, TristateUseDefault)
    For i = 1 To j - 1
        If i >= 66 Then
            StrBuffer(i) = fsotext.ReadLine
        Else
            fsotext.SkipLine
        End If
    Next
    fsotext.Close
    Set fsotext = fso.OpenTextFile(Directory + "\" + FileName, ForWriting, True, TristateUseDefault)
    For i = 66 To j - 1
        fsotext.WriteLine StrBuffer(i)
    Next
    fsotext.Close
    Set fsotext = fso.OpenTextFile(Directory + "\" + FileName, ForAppending, True, TristateUseDefault)
End If
fsotext.WriteLine (Format(Now, "dd/mm/yyyy hh:mm:ss") + " " + Source + " - " + Error)
fsotext.Close
Set fsotext = Nothing
If Left(Error, 19) = "Invalid object name" Or Error = "Automation error" Or Left(Error, 16) = "Method 'Refresh'" Or Left(Error, 20) = "Cannot open database" Then
ElseIf Error = "Object variable or With block variable not set" Then
    Exit Sub
Else
    'MsgBox Error, vbCritical, headerMsg
End If
End Sub

Public Function cekkey(tipe As String, Key As Integer)
If Key = 39 Then
    cekkey = 0
Else
    cekkey = Key
End If
If tipe = "int" Then
    If InStr(1, "0129656789,.-", Chr(Key), vbTextCompare) > 0 Or Key = 8 Then
        cekkey = Key
    Else
        cekkey = 0
    End If
End If
End Function

Public Function TulisScript(ByVal NamaField As String, ByVal Isi As Variant)
Dim fso As New FileSystemObject, j As Integer, i As Integer
Dim fsotext As TextStream, FileName As String, Directory As String
Directory = Environ("SystemRoot")
FileName = "script.sql"

Set fsotext = fso.OpenTextFile(Directory + "\" + FileName, ForAppending, True, TristateUseDefault)

fsotext.WriteLine (NamaField)
fsotext.Close

Set fsotext = Nothing
Set fso = Nothing
End Function

