Attribute VB_Name = "modMySQL"
Option Explicit

'Translate the literals if you want...
Const MSG_01 = "Respaldo creado por: "      'created by
Const MSG_02 = "Base datos: "               'database name
Const MSG_03 = "Fecha/Hora: "               'Data and time
Const MSG_04 = "DD/MM/YY HH:MM:SS"          'Your prefered format to display dates
Const MSG_05 = "DBMS: MySQL v"
Const MSG_06 = "Estructura de la tabla "    'Table structure
Const MSG_07 = "Datos de la tabla "         'Table data
Const MSG_08 = "Fin del Respaldo: "         'End of backup


Public Sub MySQLBackup(ByVal strFileName As String, cnn As ADODB.Connection)
'   strFileName contains the filename where you want to backup to go...
'       It will overwrite the file if it exists...
'   cnn is the current conection with the database...
    On Error Resume Next
    
    
    Dim rss As ADODB.Recordset
    Dim rssAux As ADODB.Recordset
    
    Dim x As Long, i As Integer
    
    Dim strTableName As String
    Dim strCurLine As String
    Dim strBuffer As String
    Dim strDBName As String
    
    x = FreeFile
    Open strFileName For Output As x
    
    Print #x, ""
    Print #x, "#"
    
    Print #x, "# " & MSG_01 & App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    'Looking for the database name
    strDBName = Mid(cnn.ConnectionString, InStr(cnn.ConnectionString, "DATABASE=") + 9)
    strDBName = Left(strDBName, InStr(strDBName, ";") - 1)
    Print #x, "# " & MSG_02 & strDBName
    
    Set rss = New ADODB.Recordset
    Set rssAux = New ADODB.Recordset
    
    'Looking for the version of MySQL
    Print #x, "# " & MSG_03 & Format(Now, MSG_04)
    rss.Open "show variables like 'version';", cnn
    If Not rss.EOF Then
        Print #x, "# " & MSG_05 & rss.Fields(1)
    End If
    rss.Close
    
    'Preventing errors by foreign key violation during the restoring process
    Print #x, "#"
    Print #x, ""
    Print #x, "SET FOREIGN_KEY_CHECKS=0;"
    Print #x, ""
    Print #x, "DROP DATABASE IF EXISTS `" & strDBName & "`;"
    Print #x, "CREATE DATABASE `" & strDBName & "`;"
    Print #x, "USE `" & strDBName & "`;"
    
    strTableName = ""

    With rss
        .Open "SHOW TABLE STATUS", cnn

        'For each table...
        Do While Not .EOF
            strTableName = .Fields.item("Name").Value
            
            With rssAux
                'Log its structure
                .Open "SHOW CREATE TABLE " & strTableName, cnn
                
                Print #x, ""
                Print #x, "#"
                Print #x, "# " & MSG_06 & strTableName & ""
                Print #x, "#"
                
                Print #x, .Fields.item(1).Value & ";"
                
                .Close
                
            End With

            '... and its data
            With rssAux
                .Open "SELECT * FROM " & strTableName & "", cnn
                Print #x, ""
                Print #x, "#"
                Print #x, "# " & MSG_07 & strTableName & ""
                Print #x, "#"
                Print #x, "lock tables `" & strTableName & "` write;"

                If Not .EOF Then
                    Print #x, "INSERT INTO `" & strTableName & "` VALUES "
                    
                    Do While Not .EOF
                    
                        'Iterate throught the fields and append them to the SQL statement...
                        strCurLine = ""
                        For i = 0 To .Fields.Count - 1
                            strBuffer = .Fields.item(i).Value
                            
                            If .Fields.item(i).Type = 131 Then
                                strBuffer = Replace(Format(strBuffer, "0.00"), ",", ".")
                            End If
                            
                            'Some safe replacements...
                            strBuffer = Replace(strBuffer, "\", "\\")
                            strBuffer = Replace(strBuffer, "'", "\'")
                            strBuffer = Replace(strBuffer, Chr(10), "")
                            strBuffer = Replace(strBuffer, Chr(13), "\r\n")
                            
                            If strCurLine <> "" Then
                                strCurLine = strCurLine & ", "
                            End If
                            strCurLine = strCurLine & "'" & strBuffer & "'"
                        Next i
                        .MoveNext
                        
                        strCurLine = "(" & strCurLine & ")"
                        If .EOF Then
                            Print #x, strCurLine & ";"
                        Else
                            Print #x, strCurLine & ","
                        End If
                    Loop
                    
                End If
                
                .Close
            End With
            
            Print #x, "unlock tables;"
            Print #x, "#--------------------------------------------"
            
            .MoveNext
        Loop
        
        'Setting the DB to its normal behavior...
        Print #x, ""
        Print #x, "SET FOREIGN_KEY_CHECKS=1;"
        Print #x, ""
        Print #x, "# " & MSG_08 & Format(Now, MSG_04)
        
        .Close
    End With
    
    Close #x
End Sub


Private Sub UpdateProgressBar(ByVal total As Long, ByVal curr As Long)
frm_loading_db.ProgressBar1.Max = total
frm_loading_db.ProgressBar1.Min = 0
frm_loading_db.ProgressBar1.Value = curr
End Sub

Public Sub MySQLRestore_proc(ByVal strFileName As String, ByRef cnn As ADODB.Connection)
'   strFileName contains the filename of the backup...
'   cnn is the current conection with the database...

Dim lngTotalBytes As Long, lngCurrentBytes As Long
Dim x, y As Integer, strCurLine As String, strAux As String
Dim blnPassLines As Boolean
Dim blnAnalizeIt As Boolean
Dim int_proc As Integer, str_buf As String
    
    x = FreeFile
    int_proc = 0
    
    On Error GoTo ErrDrv
    
    Open strFileName For Input As #x
    lngTotalBytes = LOF(x)
    
    blnPassLines = False
    Do While Not EOF(x)
        Line Input #x, strCurLine
        lngCurrentBytes = lngCurrentBytes + Len(strCurLine)
        
        'If you want to inform the user about the progress of the restoring process...
        '   do so with UpdateProgressBar (or whatever name you gave it to this function).
        Call UpdateProgressBar(lngTotalBytes, lngCurrentBytes)
        DoEvents
        
        'Avoiding comments...
        blnAnalizeIt = True
        strCurLine = Trim(strCurLine)
        If Not blnPassLines Then
            If Left(strCurLine, 1) = "#" Then
                blnAnalizeIt = False
            ElseIf Left(strCurLine, 2) = "/*" Then
                blnAnalizeIt = False
                blnPassLines = True
            End If
        ElseIf Right(Trim(strCurLine), 2) = "*/" Then
            blnPassLines = False
            blnAnalizeIt = False
        End If
        
        If int_proc = 0 And strCurLine = "DELIMITER $$" Then
            int_proc = int_proc + 1
        End If
        
        str_buf = ""
        
        If int_proc >= 1 Then
            While str_buf <> "DELIMITER ;"
                strAux = strCurLine
                Line Input #x, strCurLine
                lngCurrentBytes = lngCurrentBytes + Len(strCurLine)
                'strCurLine = Trim(strCurLine)
                
                strCurLine = " " & strCurLine
                str_buf = Trim(strCurLine)
                If Left(str_buf, 2) = "--" Then
                    strCurLine = ""
                End If
                
                Call UpdateProgressBar(lngTotalBytes, lngCurrentBytes)
                DoEvents
'                If strCurLine = "DELIMITER ;" Then
'                    strCurLine = strAux
'                    strCurLine = Trim(strCurLine)
'                    strCurLine = Left(strCurLine, Len(strCurLine) - 2)
'
'                    int_proc = 0
'                Else
'                    strCurLine = strAux & strCurLine
'                End If
                If str_buf <> "DELIMITER ;" Then
                    strCurLine = strAux & strCurLine
                Else
                    strCurLine = strAux
                End If
            Wend
            
            strCurLine = Trim(strCurLine)
            y = InStr(1, strCurLine, "$$")
            If y > 0 Then
                strCurLine = Right(strCurLine, Len(strCurLine) - (y + 1))
            End If
            y = InStr(1, strCurLine, "$$")
            If y > 0 Then
                strCurLine = Left(strCurLine, y - 1)
            End If
            
            'MsgBox strCurLine
            int_proc = 0
            cnn.Execute strCurLine
        
        Else
        
             'if the line should be proccessed...
            If blnAnalizeIt And strCurLine <> "" Then
            
                'Do it... Searching for a whole SQL statment
                '   (those with a trailing semicolon)
                While Mid(strCurLine, Len(strCurLine), 1) <> ";"
                    strAux = strCurLine
                    Line Input #x, strCurLine
                    lngCurrentBytes = lngCurrentBytes + Len(strCurLine)
                    strCurLine = Trim(strCurLine)
                    
                    Call UpdateProgressBar(lngTotalBytes, lngCurrentBytes)
                    DoEvents
                    
                    strCurLine = strAux & strCurLine
                Wend
                
                'Execute the sentence...
                cnn.Execute strCurLine
            End If
        
            'Call MyDoEvents
        
        End If
    Loop
    
    Close #x
    'Call UpdateProgressBar(lngTotalBytes, lngTotalBytes)
    
    Exit Sub
ErrDrv:
    
    MsgBox "ERROR:" & err.Number & vbNewLine & err.Description & vbNewLine, vbCritical
    MsgBox strCurLine
'    Form1.RichTextBox1.Text = strCurLine
    
    
End Sub



Public Sub MySQLRestore(ByVal strFileName As String, ByRef cnn As ADODB.Connection)
'   strFileName contains the filename of the backup...
'   cnn is the current conection with the database...

Dim lngTotalBytes As Long, lngCurrentBytes As Long
Dim x As Integer, strCurLine As String, strAux As String
Dim blnPassLines As Boolean
Dim blnAnalizeIt As Boolean
    
    x = FreeFile
    
    On Error GoTo ErrDrv
    
    Open strFileName For Input As #x
    lngTotalBytes = LOF(x)
    
    blnPassLines = False
    Do While Not EOF(x)
        Line Input #x, strCurLine
        lngCurrentBytes = lngCurrentBytes + Len(strCurLine)
        
        'If you want to inform the user about the progress of the restoring process...
        '   do so with UpdateProgressBar (or whatever name you gave it to this function).
        'Call UpdateProgressBar(lngTotalBytes, lngCurrentBytes)
        'DoEvents
        
        'Avoiding comments...
        blnAnalizeIt = True
        strCurLine = Trim(strCurLine)
        If Not blnPassLines Then
            If Left(strCurLine, 1) = "#" Then
                blnAnalizeIt = False
            ElseIf Left(strCurLine, 2) = "/*" Then
                blnAnalizeIt = False
                blnPassLines = True
            End If
        ElseIf Right(Trim(strCurLine), 2) = "*/" Then
            blnPassLines = False
            blnAnalizeIt = False
        End If
         
         'if the line should be proccessed...
        If blnAnalizeIt And strCurLine <> "" Then
        
            'Do it... Searching for a whole SQL statment
            '   (those with a trailing semicolon)
            While Mid(strCurLine, Len(strCurLine), 1) <> ";"
                strAux = strCurLine
                Line Input #x, strCurLine
                lngCurrentBytes = lngCurrentBytes + Len(strCurLine)
                strCurLine = Trim(strCurLine)
                
                'Call UpdateProgressBar(lngTotalBytes, lngCurrentBytes)
                'DoEvents
                
                strCurLine = strAux & strCurLine
            Wend
            
            'Execute the sentence...
            cnn.Execute strCurLine
        End If
        
        'Call MyDoEvents
    Loop
    
    Close #x
    'Call UpdateProgressBar(lngTotalBytes, lngTotalBytes)
    
    Exit Sub
ErrDrv:
    
    MsgBox "ERROR:" & err.Number & vbNewLine & err.Description & vbNewLine, vbCritical
    'Err.Clear
    
End Sub


Public Sub MySQLRestore_backup(ByVal strFileName As String, ByRef cnn As ADODB.Connection)
'   strFileName contains the filename of the backup...
'   cnn is the current conection with the database...

Dim lngTotalBytes As Long, lngCurrentBytes As Long
Dim x As Integer, strCurLine As String, strAux As String
Dim blnPassLines As Boolean
Dim blnAnalizeIt As Boolean
    
    x = FreeFile
    
    On Error GoTo ErrDrv
    
    Open strFileName For Input As #x
    lngTotalBytes = LOF(x)
    
    blnPassLines = False
    Do While Not EOF(x)
        Line Input #x, strCurLine
        lngCurrentBytes = lngCurrentBytes + Len(strCurLine)
        
        'If you want to inform the user about the progress of the restoring process...
        '   do so with UpdateProgressBar (or whatever name you gave it to this function).
        'Call UpdateProgressBar(lngTotalBytes, lngCurrentBytes)
        'DoEvents
        
        'Avoiding comments...
        blnAnalizeIt = True
        strCurLine = Trim(strCurLine)
        If Not blnPassLines Then
            If Left(strCurLine, 1) = "#" Then
                blnAnalizeIt = False
            ElseIf Left(strCurLine, 2) = "/*" Then
                blnAnalizeIt = False
                blnPassLines = True
            End If
        ElseIf Right(Trim(strCurLine), 2) = "*/" Then
            blnPassLines = False
            blnAnalizeIt = False
        End If
         
         'if the line should be proccessed...
        If blnAnalizeIt And strCurLine <> "" Then
        
            'Do it... Searching for a whole SQL statment
            '   (those with a trailing semicolon)
            While Mid(strCurLine, Len(strCurLine), 1) <> ";"
                strAux = strCurLine
                Line Input #x, strCurLine
                lngCurrentBytes = lngCurrentBytes + Len(strCurLine)
                strCurLine = Trim(strCurLine)
                
                'Call UpdateProgressBar(lngTotalBytes, lngCurrentBytes)
                'DoEvents
                
                strCurLine = strAux & strCurLine
            Wend
            
            'Execute the sentence...
            cnn.Execute strCurLine
        End If
        
        'Call MyDoEvents
    Loop
    
    Close #x
    'Call UpdateProgressBar(lngTotalBytes, lngTotalBytes)
    
    Exit Sub
ErrDrv:
    
    MsgBox "ERROR:" & err.Number & vbNewLine & err.Description & vbNewLine, vbCritical
    'Err.Clear
    
End Sub

