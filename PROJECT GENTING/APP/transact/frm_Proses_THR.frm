VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_proses_thr 
   Caption         =   "Proses Bonus & THR"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   7620
      TabIndex        =   13
      Top             =   630
      Width           =   2325
      Begin VB.OptionButton optBonus 
         Caption         =   "THR"
         Height          =   195
         Index           =   1
         Left            =   1350
         TabIndex        =   15
         Top             =   210
         Width           =   945
      End
      Begin VB.OptionButton optBonus 
         Caption         =   "Bonus"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   14
         Top             =   210
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin prj_genting.vbButton vbButton2 
      Height          =   345
      Left            =   12450
      TabIndex        =   12
      Top             =   1140
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Delete"
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
      MICON           =   "frm_Proses_THR.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_genting.vbButton vbButton1 
      Height          =   345
      Left            =   10980
      TabIndex        =   11
      Top             =   1140
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Proses"
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
      MICON           =   "frm_Proses_THR.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_genting.vbButton vbButton3 
      Height          =   315
      Left            =   3150
      TabIndex        =   10
      Top             =   750
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Refresh"
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
      MICON           =   "frm_Proses_THR.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cbo_religion 
      Height          =   315
      ItemData        =   "frm_Proses_THR.frx":0054
      Left            =   8520
      List            =   "frm_Proses_THR.frx":006D
      TabIndex        =   6
      Top             =   1170
      Width           =   2265
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1290
      OleObjectBlob   =   "frm_Proses_THR.frx":00B6
      TabIndex        =   2
      Top             =   1170
      Width           =   1785
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1170
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Height          =   8625
      Left            =   120
      TabIndex        =   0
      Top             =   1710
      Width           =   20025
      Begin prj_genting.LynxGrid LynxGrid1 
         Height          =   8835
         Left            =   3570
         TabIndex        =   9
         Top             =   210
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   15584
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
      Begin prj_genting.LynxGrid LynxGrid2 
         Height          =   8835
         Left            =   90
         TabIndex        =   8
         Top             =   210
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   15584
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1290
      TabIndex        =   1
      Top             =   780
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM"
      Format          =   94633987
      UpDown          =   -1  'True
      CurrentDate     =   39278
   End
   Begin VB.Label Label2 
      Caption         =   "* yyyy-MM"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1290
      TabIndex        =   17
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "PROSES BONUS / THR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4875
      TabIndex        =   16
      Top             =   0
      Width           =   3345
   End
   Begin VB.Label lbl_religion 
      AutoSize        =   -1  'True
      Caption         =   "RELIGION :"
      Height          =   195
      Left            =   7620
      TabIndex        =   7
      Top             =   1230
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal :"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   840
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Perusahaan :"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   1230
      Width           =   945
   End
End
Attribute VB_Name = "frm_proses_thr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim strsql As String

Private Sub cmbshift_Change()
    LynxGrid1.Clear
End Sub

Private Sub cbo_religion_Click()
'    Call isiGridDep
    LynxGrid1.Clear
    If TDBCombo_company.Text <> "" Then
        Call isiGridTHR
    End If
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    LynxGrid1.Clear
    LynxGrid2.Clear
    isiGridDep
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
        
    cbo_religion.ListIndex = 6
    cbo_religion.Enabled = False

    createGrid
    createGridDep
    
'    optBonus_Click (0)
'    If optBonus(0) Then
'        cbo_religion.ListIndex = 6
'        cbo_religion.Enabled = False
'    End If
    
    Call isiCompany
End Sub

Private Sub isiShift()
Dim rsshift As New ADODB.Recordset
    strsql = "select shift_code,shift_name " _
            & "from m_shift"
    rsshift.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo1.RowSource = rsshift
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "NIK", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Employee Name", 3500, , , , , , , , True
      .AddColumn "Div. Code", 1000, , , , , , , , False
      .AddColumn "Division", 2000, , , , , , , , , True
      .AddColumn "Title Code", 1000, , , , , , , , False
      .AddColumn "Title", 2000, , , , , , , , , True
      .AddColumn "Basic Income", 2000, lgAlignRightCenter, lgNumeric, "#,##", , , , , True
'      If optBonus(0).Value Then
'        .AddColumn "Bonus Value", 2000, lgAlignRightCenter, lgNumeric, "#,##", , , , , True
'      Else
        .AddColumn "Bonus / THR", 2000, lgAlignRightCenter, lgNumeric, "#,##", , , , , True
'      End If
      .AddColumn "Employee Code", 1500, lgAlignCenterCenter, , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub createGridDep()
   With LynxGrid2
      .AddColumn "Code", 800, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Department", 2500, , , , , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiCompany()
    Dim rscompany As New ADODB.Recordset
    
    strsql = "select * from m_company order by company_code"
    rscompany.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo_company.RowSource = rscompany
    'rs.Close

End Sub

Private Sub isiGridDep()
    LynxGrid2.Redraw = False
    LynxGrid2.Clear
    strsql = "select department_code,department_name " _
            & "from m_department where company_code = '" & TDBCombo_company.Text & "'"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid2.AddItem rs!DEPARTMENT_CODE & vbTab & rs!department_name
            rs.MoveNext
        Wend
    End If
    rs.Close
    LynxGrid2.Redraw = True
End Sub

Public Sub isiTHR()
Dim jmlTHR As Double

Me.MousePointer = vbHourglass

If optBonus(1).Value Then
    Dim strAgama As String
            
        If cbo_religion.ListIndex = 6 Then
            strAgama = ""
        Else
            strAgama = " AND religion = '" & cbo_religion.ListIndex & "'"
        End If
        
        '+++++++++++++++++++++++CHECK APAKAH TAHUN INI SUDAH ADA THR UNTUK DEPARTEMENT INI++++++++++
        strsql = "SELECT 1 FROM t_thr where month(tgltrans) = '" & month(DTPicker1.Value) & "' AND year(tgltrans) = '" & year(DTPicker1.Value) & "' " _
            & "AND kddivisi = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' " _
            & "AND kdcabang = '" & TDBCombo_company.Text & "' " & strAgama & " AND jenis = 0"
        rs.Open strsql, CnG, adOpenDynamic, adLockReadOnly
        If rs.RecordCount > 0 Then
            MsgBox "This Year Transactions is Already Exists..!!", vbCritical, "Warning"
            rs.Close
            Me.MousePointer = vbNormal
            Exit Sub
        End If
        rs.Close
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
        strsql = "SELECT employee_code,employee_name,a.title_code,c.title_name,a.division_code,b.division_name," _
                & "DATE(a.start_working) start_working,a.religion," _
                & "IFNULL((SELECT main_salary FROM m_salary_standard WHERE employee_code = a.employee_code ORDER BY number DESC LIMIT 1),0) basic_salary," _
                & "PERIOD_DIFF(CURDATE(),DATE(start_working)) jmlBlnKerja " _
            & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code AND a.company_code = b.company_code " _
            & "JOIN m_title c ON a.title_code = c.title_code " _
            & "WHERE a.department_code = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' AND a.company_code = '" & TDBCombo_company.Text & "' " & strAgama
        rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            CnG.BeginTrans
            rs.MoveFirst
            While Not rs.EOF
                If rs!jmlblnkerja >= 12 Then
                    jmlTHR = rs!basic_salary
                Else
                    jmlTHR = (rs!basic_salary / 12) * IIf(IsNull(rs!jmlblnkerja), 0, rs!jmlblnkerja)
                End If
                    strsql = "INSERT INTO t_thr (kdcabang,kddivisi,tgltrans,employee_code,jmlthr,jenis,tglinput,userinput,masakerja,religion) " _
                        & "VALUES " _
                        & "('" & TDBCombo_company.Text & "','" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "'," _
                        & "'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & rs!employee_code & "'," _
                        & "'" & jmlTHR & "',0,now(),'" & LOGIN_CODE & "','" & IIf(IsNull(rs!jmlblnkerja), 0, rs!jmlblnkerja) & "','" & IIf(IsNull(rs!religion), 0, rs!religion) & "')"
                    CnG.Execute strsql
                rs.MoveNext
            Wend
            CnG.CommitTrans
        End If
        rs.Close
Else
        
        '+++++++++++++++++++++++CHECK APAKAH TAHUN INI SUDAH ADA bonus UNTUK DEPARTEMENT INI++++++++++
        strsql = "SELECT 1 FROM t_thr where month(tgltrans) = '" & month(DTPicker1.Value) & "' AND year(tgltrans) = '" & year(DTPicker1.Value) & "' " _
            & "AND kddivisi = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' " _
            & "AND kdcabang = '" & TDBCombo_company.Text & "' AND jenis = 1"
        rs.Open strsql, CnG, adOpenDynamic, adLockReadOnly
        If rs.RecordCount > 0 Then
            MsgBox "This Year Transactions is Already Exists..!!", vbCritical, "Warning"
            rs.Close
            Me.MousePointer = vbNormal
            Exit Sub
        End If
        rs.Close
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
        strsql = "SELECT employee_code,employee_name,a.title_code,c.title_name,a.division_code,b.division_name," _
                & "DATE(a.start_working) start_working,a.religion," _
                & "IFNULL((SELECT main_salary FROM m_salary_standard WHERE employee_code = a.employee_code ORDER BY number DESC LIMIT 1),0) basic_salary," _
                & "IFNULL((SELECT kali_gaji FROM t_employee_performance WHERE employee_code = a.employee_code AND " _
                & "month(performance_date) = '" & month(DTPicker1.Value) & "' " _
                & "AND year(performance_date) = '" & year(DTPicker1.Value) & "' ORDER BY performance_date),0) kali_gaji," _
                & "PERIOD_DIFF(CURDATE(),DATE(start_working)) jmlBlnKerja " _
            & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code AND a.company_code = b.company_code " _
            & "JOIN m_title c ON a.title_code = c.title_code " _
            & "WHERE a.department_code = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' AND a.company_code = '" & TDBCombo_company.Text & "'"
        rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            CnG.BeginTrans
            rs.MoveFirst
            While Not rs.EOF
                    jmlTHR = rs!basic_salary * IIf(IsNull(rs!kali_gaji), 0, rs!kali_gaji)
                    
                    strsql = "INSERT INTO t_thr (kdcabang,kddivisi,tgltrans,employee_code,jmlthr,jenis,tglinput,userinput,masakerja,religion) " _
                        & "VALUES " _
                        & "('" & TDBCombo_company.Text & "','" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "'," _
                        & "'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & rs!employee_code & "'," _
                        & "'" & jmlTHR & "',1,now(),'" & LOGIN_CODE & "','" & IIf(IsNull(rs!jmlblnkerja), 0, rs!jmlblnkerja) & "','" & IIf(IsNull(rs!religion), 0, rs!religion) & "')"
                    CnG.Execute strsql
                rs.MoveNext
            Wend
            CnG.CommitTrans
        End If
        rs.Close
End If
    isiGridTHR
    Me.MousePointer = vbNormal
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute strsql
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Sub

Private Sub isiGridTHR()
Dim perAgama As Boolean
Dim strAgama As String

    If cbo_religion.Text = "..." Or cbo_religion.Text = "" Then
        MsgBox "Invalid Religion...!!", vbExclamation, "Error"
        Exit Sub
    End If
    
    If cbo_religion.ListIndex = 6 Then
        strAgama = ""
    Else
        strAgama = " AND a.religion = '" & cbo_religion.ListIndex & "'"
    End If
    
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    strsql = "SELECT a.nik,employee_name,a.title_code,c.title_name,a.division_code,b.division_name," _
            & "DATE(a.start_working) start_working," _
            & "IFNULL((SELECT main_salary FROM m_salary_standard WHERE employee_code = a.employee_code ORDER BY employee_code DESC LIMIT 1),0) basic_salary," _
            & "d.jmlthr,d.masakerja,a.employee_code " _
        & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
        & "JOIN m_title c ON a.title_code = c.title_code " _
        & "LEFT JOIN (SELECT employee_code,jmlthr,masakerja FROM t_thr " _
            & "WHERE month(tgltrans) = '" & month(DTPicker1.Value) & "' AND year(tgltrans) = '" & year(DTPicker1.Value) & "' " _
            & IIf(optBonus(0), " and jenis = 1", " and jenis = 0") & ") d on a.employee_code = d.employee_code " _
        & "WHERE a.department_code = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' AND a.company_code = '" & TDBCombo_company.Text & "' AND " _
        & "(level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))" _
        & strAgama
        
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME _
                & vbTab & rs!division_code & vbTab & rs!division_name _
                & vbTab & rs!title_code & vbTab & rs!title_name _
                & vbTab & rs!basic_salary & vbTab & rs!jmlTHR _
                & vbTab & rs!employee_code
        rs.MoveNext
        Wend
    End If
    rs.Close
    LynxGrid1.Redraw = True

End Sub

Private Sub Form_Resize()
    Frame1.Width = Me.Width - 500
    Frame1.Height = Me.Height - 1700
    LynxGrid1.Height = Frame1.Height - 400
    LynxGrid1.Width = Frame1.Width - (LynxGrid2.Width + 400)
    LynxGrid2.Height = Frame1.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_proses_thr = Nothing
End Sub

Private Sub LynxGrid2_RowColChanged()
    isiGridTHR
End Sub

Private Sub optBonus_Click(index As Integer)
If index = 0 Then
    cbo_religion.ListIndex = 6
    cbo_religion.Enabled = False
Else
    cbo_religion.Enabled = True
End If

isiGridTHR
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
        Call isiGridDep
        LynxGrid1.Clear
    End If
End Sub

Private Sub vbButton1_Click()
    isiTHR
End Sub

Private Sub vbbutton2_Click()
Dim tanya As Integer
Dim strAgama As String

    Me.MousePointer = vbHourglass
    
    If cbo_religion.ListIndex = 6 Then
        strAgama = ""
    Else
        strAgama = " AND religion = '" & cbo_religion.ListIndex & "'"
    End If
        
    tanya = MsgBox("Deleted Data Can Not be UNDO...!!" & Chr(13) & "Are You Sure To Delete This Data??", vbExclamation + vbYesNo, "Warning")
    If tanya = vbYes Then
        strsql = "DELETE FROM t_thr WHERE month(tgltrans) = '" & month(DTPicker1.Value) & "' " & _
                "AND year(tgltrans) = '" & year(DTPicker1.Value) & "' " & strAgama & _
                "AND kdcabang = '" & TDBCombo_company.Text & "'" _
                & IIf(optBonus(0), " and jenis = 1", " and jenis = 0")
                'AND kddivisi = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "'"
        CnG.Execute strsql
        isiGridTHR
    End If
    Me.MousePointer = vbNormal
End Sub

Private Sub vbButton3_Click()
    isiGridTHR
End Sub
