VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_import_employee 
   Caption         =   "IMPORT MASTER EMPLOYEE"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18030
   Icon            =   "frm_import_employee.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10785
   ScaleWidth      =   18030
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
      Left            =   960
      Top             =   630
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
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prj_tpc.vbButton vbButton2 
      Height          =   555
      Left            =   2130
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
      MICON           =   "frm_import_employee.frx":058A
      PICN            =   "frm_import_employee.frx":05A6
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
      Left            =   450
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
      MICON           =   "frm_import_employee.frx":1638
      PICN            =   "frm_import_employee.frx":1654
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
      Left            =   12930
      TabIndex        =   7
      Top             =   9390
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
      MICON           =   "frm_import_employee.frx":26E6
      PICN            =   "frm_import_employee.frx":2702
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORT MASTER EMPLOYEE"
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
      Width           =   4365
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_import_employee.frx":3794
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22710
   End
End
Attribute VB_Name = "frm_import_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset
Dim strsql As String
Dim nomer As Integer
Dim noTrans As String

Private Sub DTPicker1_Validate(Cancel As Boolean)
    LynxGrid1.Clear
End Sub

Private Sub createGrid()
   With LynxGrid1
        .AddColumn "EMP. ID NO", 1500, lgAlignCenterCenter
        .AddColumn "FULL NAME", 2500, lgAlignCenterCenter, , , , , , , True
        .AddColumn "NICK NAME", 1500, lgAlignCenterCenter, , , , , , , True
        .AddColumn "COMPANY", 1000, lgAlignCenterCenter, , , , , , , True
        .AddColumn "DEPARTMENT", 1000, lgAlignCenterCenter, , , , , , , True
        .AddColumn "SECTION", 1000, lgAlignCenterCenter, , , , , , , True
        .AddColumn "JOB TITLE", 1000, lgAlignCenterCenter, , , , , , , True
        .AddColumn "GRADE", 1000, lgAlignCenterCenter, , , , , , , True
        .AddColumn "LEVEL", 1500, lgAlignCenterCenter, , , , , , , True
        .AddColumn "STATUS", 1000, lgAlignCenterCenter, , , , , , , True
        .AddColumn "PLACE OF BIRTH", 1500, lgAlignCenterCenter, , , , , , , True
        .AddColumn "DATE OF BIRTH", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , , True
        .AddColumn "SEX", 1000, lgAlignCenterCenter, , , , , , , True
        .AddColumn "RELIGION", 1000, lgAlignCenterCenter, , , , , , , True
        .AddColumn "MARITAL STATUS", 1000, lgAlignCenterCenter, , , , , , , True
        .AddColumn "NO OF CHILDREN", 1500, lgAlignCenterCenter, , , , , , , True
        .AddColumn "NPWP", 1500, lgAlignCenterCenter, , , , , , , True
        .AddColumn "REG. DATE NPWP", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , , True
        .AddColumn "EMPLOYEE ADDRESS", 2500, lgAlignLeftCenter, , , , , , , True
        .AddColumn "NPWP ADDRESS", 2500, lgAlignLeftCenter, , , , , , , True
        .AddColumn "NO JAMSOSTEK", 1500, lgAlignLeftCenter, , , , , , , True
        .AddColumn "REG. DATE JSTK", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , , True
        .AddColumn "BANK NAME", 1000, lgAlignCenterCenter, , , , , , , True
        .AddColumn "BANK ACCOUNT", 1500, lgAlignLeftCenter, , , , , , , True
        .AddColumn "BANK ACCOUNT NAME", 2500, lgAlignLeftCenter, , , , , , , True
        .AddColumn "START WORKING", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , , True
        .AddColumn "PROBATION DATE", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , , True
        .AddColumn "PHONE NUMBER", 1500, lgAlignLeftCenter, , , , , , , True
        .AddColumn "HP NUMBER", 1500, lgAlignLeftCenter, , , , , , , True
        .AddColumn "EMAIL", 1500, lgAlignLeftCenter, , , , , , , True
        .AddColumn "NO. IDENTITY CARD", 1500, lgAlignLeftCenter, , , , , , , True
        .AddColumn "FLAG ACTIVE", 800, lgAlignLeftCenter, , , , , , , True
        .AddColumn "FLAG TRANSPORT", 800, lgAlignLeftCenter, , , , , , , True
        .AddColumn "FLAG NOT INC. LATE", 1000, lgAlignLeftCenter, , , , , , , True
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
'On Error GoTo Err
    Dim strWorksheet As String
'    Screen.MousePointer = vbHourglass
'    DoEvents
    strWorksheet = "data_employee"
    
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
                            & vbTab & .Fields(32) & vbTab & .Fields(33)
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
Dim vSex As Integer
Dim vReligion As Integer
Dim vMarital As Integer

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
                
                ProgressBar1.Value = aa
                Label1.Caption = .CellText(aa, 0) & " - " & .CellText(aa, 1)
                
                v_employee_code = Replace(.CellText(aa, 0), ".", "") & "1"
                vSex = IIf(.CellText(aa, 12) = "M", 0, 1)
                vReligion = IIf(.CellText(aa, 13) = "I", 0, IIf(.CellText(aa, 13) = "K", 1, _
                                IIf(.CellText(aa, 13) = "P", 2, IIf(.CellText(aa, 13) = "H", 3, _
                                IIf(.CellText(aa, 13) = "B", 4, 5)))))
                vMarital = IIf(.CellText(aa, 14) = "S", 0, 1)
                                                                               
                SQL = "DELETE FROM m_employee WHERE employee_code = '" & v_employee_code & "'"
                CnG.Execute SQL
                
                If .CellText(aa, 0) <> "" And .CellText(aa, 1) <> "" Then
                SQL = "INSERT INTO m_employee (employee_code,seq_no,nik,employee_name,employee_nick_name, " & _
                        "company_code,department_code,division_code,title_code,level_code," & _
                        "grade_code,status_code,place_birth,date_birth,sex,religion,marital_status," & _
                        "no_of_children,emp_address,npwp,npwp_method,npwp_registered_date,npwp_address," & _
                        "no_jamsostek,jstk_registered_date,bank_code,bank_account,bank_acc_name," & _
                        "start_working,appointment_date,phone_number,hp_number,email,country_code," & _
                        "identity_number,flag_active,flag_transport,flag_inc_late,entry_date,entry_user) " & _
                    "VALUES " & _
                        "('" & v_employee_code & "',1,'" & .CellText(aa, 0) & "','" & Replace(.CellText(aa, 1), "'", "''") & "','" & Replace(.CellText(aa, 2), "'", "''") & "'," & _
                        "'" & .CellText(aa, 3) & "','" & .CellText(aa, 4) & "','" & .CellText(aa, 5) & "','" & .CellText(aa, 6) & "','" & .CellText(aa, 8) & "'," & _
                        "'" & .CellText(aa, 7) & "','" & .CellText(aa, 9) & "','" & .CellText(aa, 10) & "'," & _
                        "'" & IIf(.CellText(aa, 11) = "", Format(Now, "yyyy-MM-dd"), Format(.CellText(aa, 11), "yyyy-MM-dd")) & "'," & _
                        "'" & vSex & "','" & vReligion & "','" & vMarital & "','" & .CellText(aa, 15) & "','" & .CellText(aa, 18) & "'," & _
                        "'" & .CellText(aa, 16) & "',0,'" & IIf(.CellText(aa, 17) = "", Format(0, "yyyy-MM-dd"), Format(.CellText(aa, 17), "yyyy-MM-dd")) & "'," & _
                        "'" & .CellText(aa, 19) & "','" & .CellText(aa, 20) & "','" & IIf(.CellText(aa, 21) = "", Format(0, "yyyy-MM-dd"), Format(.CellText(aa, 21), "yyyy-MM-dd")) & "'," & _
                        "'" & .CellText(aa, 22) & "','" & .CellText(aa, 23) & "','" & Replace(.CellText(aa, 24), "'", "''") & "'," & _
                        "'" & IIf(.CellText(aa, 25) = "", Format(Now, "yyyy-MM-dd"), Format(.CellText(aa, 25), "yyyy-MM-dd")) & "'," & _
                        "'" & IIf(.CellText(aa, 26) = "", Format(Now, "yyyy-MM-dd"), Format(.CellText(aa, 26), "yyyy-MM-dd")) & "'," & _
                        "'" & .CellText(aa, 27) & "','" & .CellText(aa, 28) & "','" & .CellText(aa, 29) & "','INA','" & .CellText(aa, 30) & "'," & _
                        "'" & .CellText(aa, 31) & "','" & .CellText(aa, 32) & "','" & .CellText(aa, 33) & "',now(),'" & LOGIN_NAME & "')"
                CnG.Execute SQL
                
                SQL = "DELETE FROM m_employee_title WHERE employee_code = '" & v_employee_code & "'"
                CnG.Execute SQL
                
                SQL = "INSERT INTO m_employee_title (employee_code,seq_no,date_title,title_code,entry_date,entry_user) " & _
                      "VALUES (" & _
                        "'" & v_employee_code & "',1," & _
                        "'" & IIf(IsNull(.CellText(aa, 25)), Format(Now, "yyyy-MM-dd"), Format(.CellText(aa, 25), "yyyy-MM-dd")) & "','" & .CellText(aa, 6) & "'," & _
                        "now(),'" & LOGIN_NAME & "')"
                CnG.Execute SQL
                
                SQL = "DELETE FROM m_employee_grade WHERE employee_code = '" & v_employee_code & "'"
                CnG.Execute SQL
                
                SQL = "INSERT INTO m_employee_grade (employee_code,seq_no,date_grade,grade_code,entry_date,entry_user) " & _
                      "VALUES (" & _
                        "'" & v_employee_code & "',1," & _
                        "'" & IIf(IsNull(.CellText(aa, 25)), Format(Now, "yyyy-MM-dd"), Format(.CellText(aa, 25), "yyyy-MM-dd")) & "','" & .CellText(aa, 7) & "'," & _
                        "now(),'" & LOGIN_NAME & "')"
                CnG.Execute SQL
        
                End If
                
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
    
    SQL = "UPDATE m_employee a JOIN td_emp_group b ON a.employee_code = b.employee_code " & _
          "SET flag_shiftable = 1"
    CnG.Execute SQL
End Sub

