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

'Public Sub FillListView2(ByRef sListView As ListView, ByRef sRecordSource As ADODB.Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean)
'Dim x As Variant
'Dim i As Byte
'Dim old_pt As Long
'
'On Error Resume Next
'    sListView.ListItems.Clear
'    If sRecordSource.RecordCount > 0 Then
'        old_pt = Screen.MousePointer
'        Screen.MousePointer = vbHourglass
'        DoEvents
'        sRecordSource.MoveFirst
'        Do While Not sRecordSource.EOF
'            If with_num = True Then
'                Set x = sListView.ListItems.Add(, , sRecordSource.AbsolutePosition, sNumIco, sNumIco)
'            Else
'                Set x = sListView.ListItems.Add(, , sRecordSource.Fields(0), sNumIco, sNumIco)
'            End If
'            For i = 1 To sNumOfFields - 1
'                If Not sRecordSource.Fields(Val(i)) = "" Then
'                    If show_first_rec = True Then
'                        x.SubItems(i) = sRecordSource.Fields(Val(i) - 1)
'                        If i = sNumOfFields - 1 Then x.SubItems(i + 1) = sRecordSource.Fields(Val(i))
'                    Else
'                        x.SubItems(i) = sRecordSource.Fields(Val(i))
'                        If i = sNumOfFields - 1 Then x.SubItems(i + 1) = sRecordSource.Fields(Val(i + 1))
'                    End If
'                End If
'            Next i
'            sRecordSource.MoveNext
'        Loop
'        i = 0
'    End If
'    Screen.MousePointer = old_pt
'    Set x = Nothing
'End Sub

'Public Sub FillCombo(ByRef combo As OsenXPComboBox, ByRef rs As Recordset)
'    Dim i As Integer
'    combo.Clear
'    With combo
'        For i = 0 To rs.Fields.Count - 1
'            .AddItem rs.Fields(i).name
'        Next
'    End With
'
'End Sub

'Public Sub FillCombobln(ByRef Combo1 As OsenXPComboBox)
'    Combo1.AddItem "Januari"
'    Combo1.AddItem "Februari"
'    Combo1.AddItem "Maret"
'    Combo1.AddItem "April"
'    Combo1.AddItem "Mei"
'    Combo1.AddItem "Juni"
'    Combo1.AddItem "Juli"
'    Combo1.AddItem "Agustus"
'    Combo1.AddItem "September"
'    Combo1.AddItem "Oktober"
'    Combo1.AddItem "November"
'    Combo1.AddItem "Desember"
'End Sub

'Procedure used to fill the LynxGrid in paging method
'Public Sub pageFillLynxGrid(ByRef sListView As LynxGrid, ByRef sRecordSource As Recordset, ByVal pos_start As Long, ByVal pos_end As Long, ByVal sNumIco As Byte, Optional warna As Boolean, Optional kolom1 As Long, Optional kolom2 As Long)
'
'Dim i As Byte, C As Long, old_pt As Long, lRow As Long, jmlkol As Integer
'Dim str  As String, lbrkol As Long
'    sListView.ImageList = MDI.ImageList2
'    sListView.Clear
'    sListView.Redraw = False
'    If sRecordSource.RecordCount < 1 Then Exit Sub
'    sRecordSource.AbsolutePosition = pos_start
'    On Error Resume Next
'    old_pt = Screen.MousePointer
'    Screen.MousePointer = vbHourglass
'    DoEvents
'    Do
'
'        str = ""
'        For i = 0 To sRecordSource.Fields.Count - 1
'            If str = "" Then
'                str = sRecordSource.Fields(CInt(i))
'            Else
'                str = str & vbTab & sRecordSource.Fields(CInt(i))
'            End If
'        Next i
'        lRow = sListView.AddItem(str)
'        sListView.RowImage(lRow) = 3
'
'        If warna = True Then
'            If sRecordSource.Fields(kolom1) <= sRecordSource.Fields(kolom2) Then
'                sListView.RowBackColor(lRow) = &H80C0FF
'                sListView.RowImage(lRow) = 5
'            End If
'        End If
'
'        If sRecordSource.AbsolutePosition >= pos_end Then
'            Exit Do
'        Else
'            sRecordSource.MoveNext
'            C = C + 1
'        End If
'    Loop
'    For jmlkol = 0 To sListView.Cols - 1
'        lbrkol = sListView.ColWidth(jmlkol)
'        sListView.ColWidthAutoSize (jmlkol)
'        If sListView.ColWidth(jmlkol) < lbrkol Then
'            sListView.ColWidth(jmlkol) = lbrkol
'        End If
'    Next
'    sListView.Redraw = True
'    Screen.MousePointer = old_pt
'    i = 0: C = 0: old_pt = 0
''    Set X = Nothing
'End Sub



