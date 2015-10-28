VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_import_salary_standard 
   Caption         =   "IMPORT DATA SALARY"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
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
      MICON           =   "frm_import_salary_standard.frx":0000
      PICN            =   "frm_import_salary_standard.frx":001C
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
      MICON           =   "frm_import_salary_standard.frx":10AE
      PICN            =   "frm_import_salary_standard.frx":10CA
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
      MICON           =   "frm_import_salary_standard.frx":215C
      PICN            =   "frm_import_salary_standard.frx":2178
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
      Picture         =   "frm_import_salary_standard.frx":320A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22710
   End
End
Attribute VB_Name = "frm_import_salary_standard"
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
        .AddColumn "SALARY DATE", 1200, lgAlignRightCenter, lgDate, "dd-MM-yyyy", , , , , , True
        .AddColumn "ID EMP. CODE", 1500, lgAlignCenterCenter
        .AddColumn "BASIC SALARY", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "FLAG POSITION", 1200, lgAlignCenterCenter, lgNumeric, , , , , , True
        .AddColumn "POSITION ALLOWANCE", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "FLAG KOPERASI", 1200, lgAlignCenterCenter, lgNumeric, , , , , , True
        .AddColumn "IURAN KOPERASI", 1200, lgAlignRightCenter, lgNumeric, "#,##"
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
    Frame1.Height = Me.Height - 3200
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
'    Screen.MousePointer = vbHourglass
'    DoEvents
    strWorksheet = "data_salary"
    
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
                            & vbTab & .Fields(6)
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
Dim SQL As String

'On Error Resume Next
    With LynxGrid1
        If .Rows > 0 Then
            ProgressBar1.Visible = True
            Label1.Visible = True
            ProgressBar1.Max = .Rows - 1
            ProgressBar1.Value = 0
            
            DoEvents
            
            For aa = 0 To .Rows - 1
                ProgressBar1.Value = aa
                Label1.Caption = .CellText(aa, 1) & " - " & .CellText(aa, 0)
                
'                strsql = "select ifnull(max(number),0) nourutdb from m_salary_standard"
'                rsnumber.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'                If rsnumber.RecordCount > 0 Then
'                    nourut = rsnumber!nourutdb + 1
'                Else
'                    nourut = 1
'                End If
'                rsnumber.Close
                          
                SQL = "SELECT * FROM m_pref_gen"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vPTKP = rscari!ptkp_code
                    vPPh = rscari!pph21_code
                    vJSTK = rscari!jamsostek_code
                    vLateTolerance = rscari!late_tolerance
                End If
                rscari.Close
                
                SQL = "SELECT * FROM m_pref_allow"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vFlagPresence = rscari!flag_presence
                    vPresenceAllow = rscari!presence_allowance
                    vFlagMeal = rscari!flag_meal
                    vMealAllow = rscari!meal_allowance
                    vFlagTransport = rscari!flag_transport
                    vTransportAllow = rscari!transport_allowance
                    vShift2Allow = rscari!shift2_allowance
                    vShift3Allow = rscari!shift3_allowance
                End If
                rscari.Close
                          
                strsql = "SELECT employee_code, company_code FROM m_employee WHERE nik = '" & .CellText(aa, 1) & "' " _
                            & "AND flag_active <> 0"
                rsnumber.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rsnumber.RecordCount > 0 Then
                    v_employee_code = rsnumber!employee_code
                    
                    SQL = "DELETE FROM m_salary_standard WHERE employee_code = '" & v_employee_code & "' " & _
                            "AND DATE(salary_date) = '" & IIf(IsNull(.CellText(aa, 0)), 0, Format(.CellText(aa, 0), "yyyy-MM-dd")) & "'"
                    CnG.Execute SQL
                    
                    If .CellText(aa, 0) <> "" And .CellText(aa, 1) <> "" Then
                    SQL = "INSERT INTO m_salary_standard (employee_code, salary_date," & _
                            "flag_basic, basic_salary, basic_salary_sunday, flag_presence,presence_allowance," & _
                            "flag_transport,transport_allowance,flag_meal,meal_allowance,flag_position,position_allowance,shift2_allowance,shift3_allowance,pph21_allowance," & _
                            "late_time_tolerance,late_amount,pph21_type,ptkp_type,jstk_type,flag_coop,coop_value,entry_date,entry_user) " & _
                            "VALUES " & _
                            "('" & v_employee_code & "','" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "'," & _
                            "1," & .CellValue(aa, 2) & ",0,'" & vFlagPresence & "'," & vPresenceAllow & "," & _
                            "'" & vFlagTransport & "'," & vTransportAllow & ",'" & vFlagMeal & "'," & vMealAllow & "," & _
                            "'" & .CellValue(aa, 3) & "'," & .CellValue(aa, 4) & "," & vShift2Allow & "," & vShift3Allow & ",0," & _
                            "" & vLateTolerance & ",0,'" & vPPh & "','" & vPTKP & "','" & vJSTK & "'," & .CellValue(aa, 5) & "," & .CellValue(aa, 6) & "," & _
                            "now(),'" & LOGIN_NAME & "')"
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

            ProgressBar1.Visible = False
            Label1.Visible = False
        End If
    End With
End Sub
