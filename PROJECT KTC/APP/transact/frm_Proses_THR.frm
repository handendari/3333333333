VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_proses_thr 
   Caption         =   "Proses THR"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   5940
      Top             =   90
   End
   Begin VB.ComboBox cbo_religion 
      Height          =   315
      ItemData        =   "frm_Proses_THR.frx":0000
      Left            =   8520
      List            =   "frm_Proses_THR.frx":0019
      TabIndex        =   10
      Top             =   600
      Width           =   2265
   End
   Begin prj_fej_jkt.vbButton vbButton1 
      Height          =   345
      Left            =   10980
      TabIndex        =   3
      Top             =   570
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Process Data"
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
      MICON           =   "frm_Proses_THR.frx":0062
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1410
      OleObjectBlob   =   "frm_Proses_THR.frx":007E
      TabIndex        =   2
      Top             =   600
      Width           =   1785
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   6
      Top             =   600
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Height          =   9315
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   20025
      Begin prj_fej_jkt.LynxGrid LynxGrid2 
         Height          =   8895
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   15690
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
      Begin prj_fej_jkt.LynxGrid LynxGrid1 
         Height          =   8835
         Left            =   3540
         TabIndex        =   5
         Top             =   210
         Width           =   12375
         _ExtentX        =   21828
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
      Left            =   810
      TabIndex        =   1
      Top             =   210
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM"
      Format          =   93126659
      UpDown          =   -1  'True
      CurrentDate     =   39278
   End
   Begin prj_fej_jkt.vbButton vbButton2 
      Height          =   345
      Left            =   12600
      TabIndex        =   4
      Top             =   570
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Delete Data"
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
      MICON           =   "frm_Proses_THR.frx":1FE4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_fej_jkt.vbButton vbButton3 
      Height          =   345
      Left            =   2610
      TabIndex        =   12
      Top             =   180
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
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
      MICON           =   "frm_Proses_THR.frx":2000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lbl_religion 
      AutoSize        =   -1  'True
      Caption         =   "RELIGION :"
      Height          =   195
      Left            =   7620
      TabIndex        =   11
      Top             =   660
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Date :"
      Height          =   195
      Left            =   300
      TabIndex        =   8
      Top             =   270
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Branch Office :"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   660
      Width           =   1065
   End
End
Attribute VB_Name = "frm_proses_thr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim strsql As String
Dim access As String
Dim rscompany As New ADODB.Recordset

Private Sub cmbshift_Change()
    LynxGrid1.Clear
End Sub

Private Sub cbo_religion_Click()
'    Call isiGridDep
    LynxGrid1.Clear
    Call isiGridTHR
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    LynxGrid1.Clear
    LynxGrid2.Clear
    isiGridDep
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    
    createGrid
    createGridDep
    
    Call isiCompany
    
    Timer1.Enabled = True
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
      .AddColumn "Employee Name", 3000, , , , , , , , True
      .AddColumn "Start Working", 1500, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , True
      .AddColumn "Div. Code", 1000, , , , , , , , False
      .AddColumn "Division", 1800, , , , , , , , , True
      .AddColumn "Title Code", 1000, , , , , , , , False
      .AddColumn "Title", 2000, , , , , , , , , True
      .AddColumn "Basic Income", 1500, lgAlignRightCenter, lgNumeric, "#,##", , , , , True
      .AddColumn "THR Value", 1500, lgAlignRightCenter, lgNumeric, "#,##", , , , , True
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
            LynxGrid2.AddItem rs!department_code & vbTab & rs!department_name
            rs.MoveNext
        Wend
    End If
    rs.Close
    LynxGrid2.Redraw = True
End Sub

Public Sub isiTHR()
Dim jmlTHR As Double
Dim strAgama As String
Dim tglProses As String
Dim jmlBulan As Double
Dim blnTHR As String

    Me.MousePointer = vbHourglass
    
    If cbo_religion.ListIndex = 6 Then
        strAgama = ""
    Else
        strAgama = " AND religion = '" & cbo_religion.ListIndex & "'"
    End If
    
    '+++++++++++++++++++++++CHECK APAKAH TAHUN INI SUDAH ADA THR UNTUK DEPARTEMENT INI++++++++++
    strsql = "SELECT 1 FROM t_thr where month(tgltrans) = '" & Month(DTPicker1.Value) & "' AND year(tgltrans) = '" & Year(DTPicker1.Value) & "' " _
        & "AND kddivisi = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' " _
        & "AND kdcabang = '" & TDBCombo_company.Text & "' " & strAgama
    rs.Open strsql, CnG, adOpenDynamic, adLockReadOnly
    If rs.RecordCount > 0 Then
        MsgBox "This Year Transactions is Already Exists..!!", vbCritical, "Warning"
        rs.Close
        Me.MousePointer = vbNormal
        Exit Sub
    End If
    rs.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
'    access = IIf(LOGIN_LEVEL = 100, "", "AND (a.managerial_access = 0 OR a.managerial_access IS NULL)")
    
'    strsql = "SELECT employee_code,employee_name,a.title_code,c.title_name,a.division_code,b.division_name," _
'            & "DATE(a.start_working) start_working,a.religion," _
'            & "IFNULL((SELECT salary FROM m_salary WHERE employee_code = a.employee_code ORDER BY salary_date DESC LIMIT 1),0) basic_salary," _
'            & "PERIOD_DIFF(Now(),DATE(start_working)) jmlBlnKerja " _
'        & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code AND a.company_code = b.company_code " _
'        & "JOIN m_title c ON a.title_code = c.title_code " _
'        & "WHERE a.department_code = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' AND a.company_code = '" & TDBCombo_company.Text & "' " & strAgama
'    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    strsql = "SELECT employee_code,employee_name,a.title_code,c.title_name,a.division_code,b.division_name," _
            & "DATE(a.start_working) start_working,a.religion," _
            & "IFNULL((SELECT salary FROM m_salary WHERE employee_code = a.employee_code ORDER BY salary_date DESC LIMIT 1),0) basic_salary," _
            & "DATE(start_working) start_working " _
        & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code AND a.company_code = b.company_code " _
        & "JOIN m_title c ON a.title_code = c.title_code " _
        & "WHERE a.department_code = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' AND a.company_code = '" & TDBCombo_company.Text & "' " & strAgama
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        CnG.BeginTrans
        rs.MoveFirst
        While Not rs.EOF
            blnTHR = DateAdd("m", 1, DTPicker1.Value)
            
            If Month(DTPicker1) < 10 Then
                tglProses = Year(blnTHR) & "-0" & Month(blnTHR) & "-" & getEndDay(Month(blnTHR), Year(blnTHR))
            Else
                tglProses = Year(blnTHR) & "-" & Month(blnTHR) & "-" & getEndDay(Month(blnTHR), Year(blnTHR))
            End If
            
            jmlBulan = DateDiff("m", rs!start_working, tglProses)
            
            If jmlBulan >= 12 Then
                jmlTHR = rs!basic_salary
            Else
                jmlTHR = (rs!basic_salary / 12) * jmlBulan
            End If
                strsql = "INSERT INTO t_thr (kdcabang,kddivisi,tgltrans,employee_code,jmlthr,jenis,tglinput,userinput,masakerja,religion) " _
                    & "VALUES " _
                    & "('" & TDBCombo_company.Text & "','" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "'," _
                    & "'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & rs!employee_code & "'," _
                    & "'" & jmlTHR & "',0,now(),'" & LOGIN_CODE & "','" & jmlBulan & "','" & IIf(IsNull(rs!religion), 0, rs!religion) & "')"
                CnG.Execute strsql
            rs.MoveNext
        Wend
        CnG.CommitTrans
    End If
    rs.Close
    isiGridTHR
    Me.MousePointer = vbNormal
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
    
    access = IIf(LOGIN_LEVEL = 100, "", "AND (a.managerial_access = 0 OR a.managerial_access IS NULL)")
    
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    strsql = "SELECT a.employee_code,employee_name,a.title_code,c.title_name,a.division_code,b.division_name," _
            & "DATE(a.start_working) start_working," _
            & "IFNULL((SELECT salary FROM m_salary WHERE employee_code = a.employee_code ORDER BY salary_date DESC LIMIT 1),0) basic_salary," _
            & "d.jmlthr,d.masakerja " _
        & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
        & "JOIN m_title c ON a.title_code = c.title_code " _
        & "LEFT JOIN (SELECT employee_code,jmlthr,masakerja FROM t_thr " _
            & "WHERE month(tgltrans) = '" & Month(DTPicker1.Value) & "' AND year(tgltrans) = '" & Year(DTPicker1.Value) & "') d on a.employee_code = d.employee_code " _
        & "WHERE a.department_code = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' " & access & " AND a.company_code = '" & TDBCombo_company.Text & "' " & strAgama
        
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!employee_code & vbTab & rs!employee_name & vbTab & rs!start_working _
                & vbTab & rs!division_code & vbTab & rs!division_name _
                & vbTab & rs!title_code & vbTab & rs!title_name _
                & vbTab & rs!basic_salary & vbTab & rs!jmlTHR
        rs.MoveNext
        Wend
    End If
    rs.Close
    LynxGrid1.Redraw = True

End Sub

Private Sub Form_Resize()
On Error Resume Next
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

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
        Call isiGridDep
        LynxGrid1.Clear
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call set_company_mode_rs(rscompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub vbButton1_Click()
    isiTHR
End Sub

Private Sub vbButton2_Click()
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
        strsql = "DELETE FROM t_thr WHERE month(tgltrans) = '" & Month(DTPicker1.Value) & "' " & _
                "AND year(tgltrans) = '" & Year(DTPicker1.Value) & "' " & strAgama & _
                "AND kdcabang = '" & TDBCombo_company.Text & "'"
                'AND kddivisi = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "'"
        CnG.Execute strsql
        isiGridTHR
    End If
    Me.MousePointer = vbNormal
End Sub

Private Sub vbButton3_Click()
    isiGridTHR
End Sub
