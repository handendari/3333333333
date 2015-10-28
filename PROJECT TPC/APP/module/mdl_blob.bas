Attribute VB_Name = "modBlob"
Option Explicit

Public com As ADODB.Command
Global strcon As New ADODB.Connection
Public Const CHUNK_SIZE     As Long = 16384

Dim rsImage                 As ADODB.Recordset

Dim i                       As Long
Dim lsize                   As Long
Dim iChunks                 As Long
Dim nFragmentOffset         As Long
Dim lchunks                 As Long

Dim nHandle                 As Integer
Dim varChunk()              As Byte

Public Function fileExists(ByVal strNamaFile As String) As Boolean
    If Not (Len(strNamaFile) > 0) Then fileExists = False: Exit Function

    If Dir$(strNamaFile, vbNormal) = "" Then
        fileExists = False
    Else
        fileExists = True
    End If
End Function

Public Sub closeRecordset(ByVal vRs As ADODB.Recordset)
    On Error Resume Next

    If Not (vRs Is Nothing) Then
        If vRs.State = adStateOpen Then
            vRs.Close
            Set vRs = Nothing
        End If
    End If
End Sub

Public Function addImageToDB(ByVal query As String, ByVal imageName As String, ByVal imageField As String) As Boolean
'    On Error GoTo errHandle

    Set rsImage = New ADODB.Recordset
    rsImage.Open query, CnG, adOpenKeyset, adLockOptimistic
    If Not rsImage.EOF Then
        nHandle = FreeFile
        Open imageName For Binary Access Read As nHandle
        lsize = LOF(nHandle)
        If nHandle = 0 Then Close nHandle

        lchunks = lsize / CHUNK_SIZE
        nFragmentOffset = lsize Mod CHUNK_SIZE

        ReDim varChunk(nFragmentOffset)
        Get nHandle, , varChunk()
        rsImage(imageField).AppendChunk varChunk()

        ReDim varChunk(CHUNK_SIZE)
        For i = 1 To lchunks
            Get nHandle, , varChunk()
            rsImage(imageField).AppendChunk varChunk()
            DoEvents
        Next
        rsImage.Update
    End If
    Call closeRecordset(rsImage)

    addImageToDB = True

    Exit Function
errHandle:
    addImageToDB = False
End Function

Public Function getImageFromDB(ByVal query As String) As IPictureDisp
    Dim sFile           As String

    On Error GoTo errHandle

    Set rsImage = New ADODB.Recordset
    rsImage.Open query, CnG, adOpenForwardOnly, adLockReadOnly
    If Not rsImage.EOF Then
        If Not IsNull(rsImage(0).Value) Then
            nHandle = FreeFile

            sFile = App.Path & "\output.bin"
            If fileExists(sFile) Then Kill sFile
            DoEvents

            Open sFile For Binary Access Write As nHandle

            lsize = rsImage(0).ActualSize
            iChunks = lsize \ CHUNK_SIZE
            nFragmentOffset = lsize Mod CHUNK_SIZE

            varChunk() = rsImage(0).GetChunk(nFragmentOffset)
            Put nHandle, , varChunk()
            For i = 1 To iChunks
                 ReDim varChunk(CHUNK_SIZE) As Byte

                 varChunk() = rsImage(0).GetChunk(CHUNK_SIZE)
                 Put nHandle, , varChunk()
                 DoEvents
            Next
            Close nHandle

            Set getImageFromDB = LoadPicture(sFile, , vbLPColor)

        Else
            Set getImageFromDB = Nothing
        End If

    Else
        Set getImageFromDB = Nothing
    End If
    Call closeRecordset(rsImage)

    Exit Function
    
errHandle:
    Set getImageFromDB = Nothing
End Function
