VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_import_salary_manual 
   Caption         =   "IMPORT DATA SALARY"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form2"
   ScaleHeight     =   10785
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   5100
      TabIndex        =   1
      Top             =   9480
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   900
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   7980
      Left            =   120
      TabIndex        =   0
      Top             =   1050
      Width           =   20025
      Begin prj_tpc.LynxGrid LynxGrid1 
         Height          =   7635
         Left            =   120
         TabIndex        =   3
         Top             =   210
         Width           =   13665
         _ExtentX        =   24104
         _ExtentY        =   13467
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorSel    =   12937777
         ForeColorSel    =   16777215
         CustomColorFrom =   16572875
         CustomColorTo   =   14722429
         GridColor       =   16367254
         FocusRectColor  =   9895934
         Appearance      =   0
         ColumnHeaderSmall=   0   'False
         TotalsLineShow  =   0   'False
         FocusRowHighlightKeepTextForecolor=   0   'False
         ShowRowNumbers  =   0   'False
         ShowRowNumbersVary=   0   'False
         AllowColumnResizing=   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   30
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prj_tpc.vbButton vbButton2 
      Height          =   555
      Left            =   2280
      TabIndex        =   5
      Top             =   9480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   979
      BTYPE           =   14
      TX              =   "&Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_import_salary_manual.frx":0000
      PICN            =   "frm_import_salary_manual.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton vbButton1 
      Height          =   555
      Left            =   600
      TabIndex        =   6
      Top             =   9480
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   979
      BTYPE           =   14
      TX              =   "&Browse"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_import_salary_manual.frx":10AE
      PICN            =   "frm_import_salary_manual.frx":10CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   12960
      TabIndex        =   7
      Top             =   9480
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1244
      BTYPE           =   14
      TX              =   "&Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_import_salary_manual.frx":215C
      PICN            =   "frm_import_salary_manual.frx":2178
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORT DATA SALARY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      TabIndex        =   4
      Top             =   150
      Width           =   6345
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   5130
      TabIndex        =   2
      Top             =   9780
      Visible         =   0   'False
      Width           =   7275
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_import_salary_manual.frx":320A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22710
   End
End
Attribute VB_Name = "frm_import_salary_manual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset
Dim strsql As String

Dim vPTKP As String, vPPh As String, vJSTK As String, vLateTolerance As Double
Dim vFlagPresence As Integer, vFlagMeal As Integer, vFlagTransport As Integer
Dim vPresenceAllow As Double, vMealAllow As Double, vTransportAllow As Double
Dim vShift2Allow As Double, vShift3Allow As Double

Private Sub DTPicker1_Validate(Cancel As Boolean)
    LynxGrid1.Clear
End Sub

Private Sub createGrid()
   With LynxGrid1
        .AddColumn "PERIODE", 1200, lgAlignCenterCenter, , , , , , , True               '0
        .AddColumn "ID EMP. CODE", 1200, lgAlignCenterCenter, , , , , , , True          '1
        .AddColumn "BASIC SALARY", 1200, lgAlignRightCenter, lgNumeric, "#,##.##"          '2
        .AddColumn "JK JKK", 1200, lgAlignRightCenter, lgNumeric, "#,##.##"                '3
        .AddColumn "LEAVE (Day)", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"           '4
        .AddColumn "PL (Hour)", 1000, lgAlignRightCenter, lgNumeric, "#,##.###"             '5
        .AddColumn "SL (Hour)", 1000, lgAlignRightCenter, lgNumeric, "#,##.###"             '6
        .AddColumn "AL (Hour)", 1000, lgAlignRightCenter, lgNumeric, "#,##.###"             '7
        .AddColumn "LATE (Hour)", 1000, lgAlignRightCenter, lgNumeric, "#,##.###"           '8
        .AddColumn "LATE FREQ. (Times)", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"    '9
        .AddColumn "OT X1.5 WD", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"            '10
        .AddColumn "OT X2 WD", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"              '11
        .AddColumn "OT X2 HOL", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"             '12
        .AddColumn "OT X3 HOL", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"             '13
        .AddColumn "OT X4 HOL", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"             '14
        .AddColumn "OT X4 SPEC.", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"           '15
        .AddColumn "OT X6 SPEC.", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"           '16
        .AddColumn "PL EXPENSE", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"            '17
        .AddColumn "PL NOT APPROVE EXPENSE", 1000, lgAlignRightCenter, lgNumeric, "#,##.##" '18
        .AddColumn "ALPHA EXPENSE", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"         '19
        .AddColumn "LATE EXPENSE", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"          '20
        .AddColumn "OVERTIME", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"              '21
        .AddColumn "ATTENDANCE ALLOW.", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"     '22
        .AddColumn "POSITION ALLOW.", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"       '23
        .AddColumn "OTHER ALLOW.", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"          '24
        .AddColumn "MEAL DAYS", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"             '25
        .AddColumn "MEAL ALLOW", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"            '26
        .AddColumn "TRANS DAYS", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"            '27
        .AddColumn "TRANS ALLOW", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"           '28
        .AddColumn "SHIFT A DAYS", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"          '29
        .AddColumn "SHIFT A ALLOW", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"         '30
        .AddColumn "SHIFT N DAYS", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"          '31
        .AddColumn "SHIFT N ALLOW", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"         '32
        .AddColumn "REMARKS", 1200, lgAlignLeftCenter, , , , , , , True                 '33
        .AddColumn "TOT. RECEIVED", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"         '34
        .AddColumn "JAMSOSTEK", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"             '35
        .AddColumn "INCOME TAX", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"            '36
        .AddColumn "ACT. RECEIVED", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"         '37
        .AddColumn "TAX CORRECTION", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"        '38
        .AddColumn "COOP. DUES", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"            '39
        .AddColumn "COOP. INST.", 1000, lgAlignRightCenter, lgNumeric, "#,##.##"            '40
        .BackColorBkg = &HFCE1CB
        .Redraw = True
   End With
    
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    createGrid
End Sub

Private Sub Form_Resize()
    Frame1.Width = Me.Width - 500
    Frame1.Height = Me.Height - 3500
    LynxGrid1.Height = Frame1.Height - 400
    LynxGrid1.Width = Frame1.Width - 400
'    vbButton1.Top = LynxGrid1.Top + LynxGrid1.Height + 100
'    vbButton2.Top = LynxGrid1.Top + LynxGrid1.Height + 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_import_salary = Nothing
End Sub

Private Sub vbButton1_Click()
    CommonDialog1.Filter = "XLS|*.xls"
    CommonDialog1.initDir = App.Path
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
        Call fill_grid_excel_m(CommonDialog1.FileName)
    End If
End Sub

Private Sub fill_grid_excel_m(str_file_name As String)
On Error GoTo Err
    Dim strWorksheet As String
    strWorksheet = "data_salary_manual"
    
    Adodc1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & str_file_name & ";Extended Properties=Excel 8.0"
    
    Adodc1.RecordSource = "select * from [" & strWorksheet & "$]"
    Adodc1.Refresh
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    With Adodc1.Recordset
        If .RecordCount > 0 Then
            Me.MousePointer = vbHourglass
            .MoveFirst
            While Not .EOF
                'If Adodc1.Recordset!employee_code <> "" Or Adodc1.Recordset!employee_code Is Null Then
                    LynxGrid1.AddItem .Fields(0) & vbTab & .Fields(1) _
                            & vbTab & .Fields(2) & vbTab & .Fields(3) _
                            & vbTab & .Fields(4) & vbTab & .Fields(5) _
                            & vbTab & .Fields(6) & vbTab & .Fields(7) _
                            & vbTab & .Fields(8) & vbTab & .Fields(9) _
                            & vbTab & .Fields(10) & vbTab & .Fields(11) _
                            & vbTab & .Fields(12) & vbTab & .Fields(13) _
                            & vbTab & .Fields(14) & vbTab & .Fields(15) _
                            & vbTab & .Fields(16) & vbTab & .Fields(17) _
                            & vbTab & .Fields(18) & vbTab & .Fields(19) _
                            & vbTab & .Fields(20) & vbTab & .Fields(21) _
                            & vbTab & .Fields(22) & vbTab & .Fields(23) _
                            & vbTab & .Fields(24) & vbTab & .Fields(25) _
                            & vbTab & .Fields(26) & vbTab & .Fields(27) _
                            & vbTab & .Fields(28) & vbTab & .Fields(29) _
                            & vbTab & .Fields(30) & vbTab & .Fields(31) _
                            & vbTab & .Fields(32) & vbTab & .Fields(33) _
                            & vbTab & .Fields(34) & vbTab & .Fields(35) _
                            & vbTab & .Fields(36) & vbTab & .Fields(37) _
                            & vbTab & .Fields(38) & vbTab & .Fields(39) _
                            & vbTab & .Fields(40)
                'End If
                .MoveNext
            Wend
            Me.MousePointer = vbNormal
        End If
    End With
    LynxGrid1.Redraw = True
    Exit Sub
    
Err:
    MsgBox Err.Description, vbExclamation, "Message Error!"
    Exit Sub
End Sub

Private Sub vbbutton2_Click()
Dim aa As Integer
Dim rsnumber As New ADODB.Recordset
Dim nourut As Long
Dim v_employee_code As String
Dim v_company_code As String, v_company_name As String
Dim periode As String, tgl1 As String, tgl2 As String
Dim vTotGajiBersih As Double
Dim vGajiTunj As Double
Dim SQL As String

'On Error Resume Next
    With LynxGrid1
        If .Rows > 0 Then
            ProgressBar1.Visible = True
            Label1.Visible = True
            ProgressBar1.Max = .Rows
            ProgressBar1.Value = 0
            
            DoEvents
            
            For aa = 0 To .Rows - 1
                ProgressBar1.Value = aa + 1
                Label1.Caption = .CellText(aa, 1) & " - " & .CellText(aa, 0)
                          
                strsql = "SELECT a.employee_code, a.company_code, b.company_name " & _
                         "FROM m_employee a JOIN m_company b ON a.company_code = b.company_code " & _
                         "WHERE a.nik = '" & .CellText(aa, 1) & "' " _
                            & "AND a.flag_active <> 0"
                rsnumber.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rsnumber.RecordCount > 0 Then
                    v_employee_code = rsnumber!employee_code
                    v_company_code = rsnumber!COMPANY_CODE
                    v_company_name = rsnumber!company_name
                    
                    periode = .CellText(aa, 0) & "-01"
                    tgl1 = Format(DateAdd("m", -1, periode), "yyyy-MM-") & "21"
                    tgl2 = Format(periode, "yyyy-MM-") & "20"
                    
                    If .CellText(aa, 0) <> "" And .CellText(aa, 1) <> "" Then
                        SQL = "DELETE FROM h_salary_manual " & _
                              "WHERE employee_code = '" & v_employee_code & "' " & _
                                "AND month = '" & .CellText(aa, 0) & "'"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary_manual " & _
                              "VALUES (" & _
                                "'" & .CellText(aa, 0) & "','" & v_employee_code & "'," & .CellValue(aa, 4) & "," & _
                                "" & .CellValue(aa, 5) & "," & .CellValue(aa, 6) & "," & .CellValue(aa, 7) & "," & _
                                "" & .CellValue(aa, 8) & "," & .CellValue(aa, 9) & "," & .CellValue(aa, 10) & "," & _
                                "" & .CellValue(aa, 11) & "," & .CellValue(aa, 12) & "," & .CellValue(aa, 13) & "," & _
                                "" & .CellValue(aa, 14) & "," & .CellValue(aa, 15) & "," & .CellValue(aa, 16) & "," & _
                                "" & .CellValue(aa, 22) & "," & .CellValue(aa, 24) & "," & .CellValue(aa, 25) & "," & _
                                "" & .CellValue(aa, 27) & "," & .CellValue(aa, 29) & "," & .CellValue(aa, 31) & "," & _
                                "" & .CellValue(aa, 38) & "," & .CellValue(aa, 39) & "," & .CellValue(aa, 40) & "," & _
                                "'" & Replace(.CellText(aa, 33), "'", "''") & "',Now(),'" & LOGIN_NAME & "' " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "DELETE FROM h_salary " & _
                              "WHERE employee_code = '" & v_employee_code & "' " & _
                                "AND month = '" & .CellText(aa, 0) & "'"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-01'," & _
                                "'" & v_company_code & "','BASIC SALARY', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 2) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-034'," & _
                                "'" & v_company_code & "','JSTK ACCIDENT', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 3) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-035'," & _
                                "'" & v_company_code & "','JSTK LIFE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "0,1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-0711'," & _
                                "'" & v_company_code & "','PL EXPENSE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 17) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-0712'," & _
                                "'" & v_company_code & "','PL NOT APPROVED EXPENSE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 18) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-0713'," & _
                                "'" & v_company_code & "','ALPHA EXPENCE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 19) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-077'," & _
                                "'" & v_company_code & "','LATE EXPENCE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 20) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-025'," & _
                                "'" & v_company_code & "','OVERTIME', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 21) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-020'," & _
                                "'" & v_company_code & "','PRESENCE ALLOWANCE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 22) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-021'," & _
                                "'" & v_company_code & "','POSITION ALLOWANCE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 23) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-028'," & _
                                "'" & v_company_code & "','POSITION ALLOWANCE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 24) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-023'," & _
                                "'" & v_company_code & "','MEAL ALLOWANCE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 26) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-024'," & _
                                "'" & v_company_code & "','TRANSPORT ALLOWANCE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 28) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-026'," & _
                                "'" & v_company_code & "','SHIFT2 ALLOWANCE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 30) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-027'," & _
                                "'" & v_company_code & "','SHIFT3 ALLOWANCE', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 32) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        vTotGajiBersih = (.CellValue(aa, 34) + .CellValue(aa, 3))
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-28'," & _
                                "'" & v_company_code & "','TOT. GAJI BERSIH', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & vTotGajiBersih & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        vGajiTunj = (.CellValue(aa, 2) + .CellValue(aa, 3) + .CellValue(aa, 21) + _
                                     .CellValue(aa, 22) + .CellValue(aa, 23) + .CellValue(aa, 24) + _
                                     .CellValue(aa, 26) + .CellValue(aa, 28) + .CellValue(aa, 30) + _
                                     .CellValue(aa, 32))
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-06'," & _
                                "'" & v_company_code & "','TOT. GAJI + TUNJ', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & vGajiTunj & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-281'," & _
                                "'" & v_company_code & "','JAMSOSTEK', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 35) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-36'," & _
                                "'" & v_company_code & "','PPH21', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 36) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-37'," & _
                                "'" & v_company_code & "','SISA GAJI', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 37) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-037'," & _
                                "'" & v_company_code & "','TAX CORRECTION', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 38) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-076'," & _
                                "'" & v_company_code & "','IURAN KOPERASI', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 39) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-075'," & _
                                "'" & v_company_code & "','CICILAN KOPERASI', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "" & .CellValue(aa, 40) & ",1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_salary (month, employee_code, salary_code, company_code," & _
                                    "salary_name, date_from, date_to, salary_value, flag_manual) " & _
                              "VALUES ( " & _
                                "'" & .CellText(aa, 0) & "', '" & v_employee_code & "', 'SU-074'," & _
                                "'" & v_company_code & "','CICILAN PERUSAHAAN', '" & tgl1 & "', '" & tgl2 & "'," & _
                                "0,1 " & _
                              ")"
                        CnG.Execute SQL
                        
                        
                        SQL = "DELETE FROM h_d_salary WHERE company_code = '" & v_company_code & "' and left(month,7) = '" & .CellText(aa, 0) & "'"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_d_salary (month,periode_from,periode_to,company_code,company_name) " & _
                            "VALUES " & _
                            "('" & periode & "','" & tgl1 & "','" & tgl2 & "'," & _
                            "'" & v_company_code & "','" & Replace(v_company_name, "'", "''") & "')"
                        CnG.Execute SQL
                    End If
                End If
                rsnumber.Close
                
                
                DoEvents
            Next
            MsgBox "Import Data Success...!!!"
            
            '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
            strsql = "Update temp_sal_proses set salary_proses = 0"
            CnG.Execute strsql
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            Call frm_trans_salary_process.load_data

            ProgressBar1.Visible = False
            Label1.Visible = False
        End If
    End With
End Sub


