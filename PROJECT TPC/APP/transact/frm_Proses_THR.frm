VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_proses_thr 
   Caption         =   "THR PROCCESS"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   Icon            =   "frm_Proses_THR.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Tax Method"
      Height          =   525
      Left            =   11310
      TabIndex        =   13
      Top             =   930
      Width           =   2835
      Begin VB.OptionButton optTax 
         Caption         =   "New"
         Height          =   195
         Index           =   1
         Left            =   1290
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optTax 
         Caption         =   "Old"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   0
      Top             =   0
   End
   Begin VB.ComboBox cbo_religion 
      Height          =   315
      ItemData        =   "frm_Proses_THR.frx":058A
      Left            =   8520
      List            =   "frm_Proses_THR.frx":05A3
      TabIndex        =   6
      Top             =   1590
      Width           =   2265
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1230
      OleObjectBlob   =   "frm_Proses_THR.frx":05EC
      TabIndex        =   2
      Top             =   1590
      Width           =   1695
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3030
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1590
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Height          =   9315
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   20025
      Begin prj_tpc.LynxGrid LynxGrid1 
         Height          =   8325
         Left            =   60
         TabIndex        =   8
         Top             =   150
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   14684
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
      Left            =   1230
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   97452035
      UpDown          =   -1  'True
      CurrentDate     =   39278
   End
   Begin prj_tpc.vbButton vbButton3 
      Height          =   495
      Left            =   3030
      TabIndex        =   9
      Top             =   960
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   873
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
      MICON           =   "frm_Proses_THR.frx":2552
      PICN            =   "frm_Proses_THR.frx":256E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton vbButton2 
      Height          =   465
      Left            =   12750
      TabIndex        =   10
      Top             =   1500
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&Delete"
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
      MICON           =   "frm_Proses_THR.frx":3600
      PICN            =   "frm_Proses_THR.frx":361C
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
      Height          =   465
      Left            =   11310
      TabIndex        =   11
      Top             =   1500
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&Proccess"
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
      MICON           =   "frm_Proses_THR.frx":46AE
      PICN            =   "frm_Proses_THR.frx":46CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "THR PROCCESS"
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
      Left            =   390
      TabIndex        =   12
      Top             =   150
      Width           =   4845
   End
   Begin VB.Label lbl_religion 
      AutoSize        =   -1  'True
      Caption         =   "RELIGION :"
      Height          =   195
      Left            =   7620
      TabIndex        =   7
      Top             =   1650
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "DATE :"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   1140
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY :"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   1650
      Width           =   885
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_Proses_THR.frx":575C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   24120
   End
End
Attribute VB_Name = "frm_proses_thr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCompany As New ADODB.Recordset
    
Dim rs As New ADODB.Recordset
Dim access As String
Dim strAgama As String
Dim vExist As Boolean

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
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    
    createGrid
    timer1.Enabled = True
    
    Call isiCompany
End Sub

Private Sub isiShift()
Dim rsShift As New ADODB.Recordset
    SQL = "select shift_code,shift_name " _
            & "from m_shift"
    rsShift.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo1.RowSource = rsShift
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "EMP. ID NO.", 1000, lgAlignCenterCenter, , , , , , , True
      .AddColumn "EMPLOYEE NAME", 3000, , , , , , , , True
      .AddColumn "START WORKING", 1500, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , True
      .AddColumn "Dept. Code", 1000, , , , , , , , False
      .AddColumn "DEPARTMENT", 1800, , , , , , , , , True
      .AddColumn "Title Code", 1000, , , , , , , , False
      .AddColumn "JOB TITLE", 2000, , , , , , , , , True
      .AddColumn "BASIC SALARY", 1500, lgAlignRightCenter, lgNumeric, "#,##", , , , , True
      .AddColumn "MASA KERJA", 1000, lgAlignCenterCenter, , , , , , , True
      .AddColumn "THR VALUE", 1500, lgAlignRightCenter, lgNumeric, "#,##", , , , , True
      .AddColumn "THR TAX", 1500, lgAlignRightCenter, lgNumeric, "#,##", , , , , True
      .AddColumn "ACTUAL", 1500, lgAlignRightCenter, lgNumeric, "#,##", , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiCompany()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo_company.RowSource = rsCompany
    'rs.Close

End Sub

Private Sub isiTHR()
Dim rs2 As New ADODB.Recordset

Dim jmlTHR As Double
Dim tglProses As String
Dim jmlBulan As Double
Dim blnTHR As String
Dim vUMK As Double
Dim vFlagHL As Integer
Dim vTHR_HL As Double
Dim vFlag_percentage As Integer
Dim vStartWorking As String

    Me.MousePointer = vbHourglass
    vExist = False
    
    If cbo_religion.Text = "..." Or cbo_religion.Text = "" Then
        MsgBox "Invalid Religion...!!", vbExclamation, "Error"
        Exit Sub
    End If
    
    SQL = "SELECT * FROM m_umk WHERE date(umk_date) <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
            "ORDER BY umk_date DESC LIMIT 1"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vUMK = rs!umk_value
    End If
    rs.Close
    
    If cbo_religion.ListIndex = 6 Then
        strAgama = ""
    Else
        strAgama = " AND religion = '" & cbo_religion.ListIndex & "'"
    End If
    
    '+++++++++++++++++++++++CHECK APAKAH TAHUN INI SUDAH ADA THR UNTUK DEPARTEMENT INI++++++++++
    Dim vtglTrans As String
    SQL = "SELECT tgltrans FROM t_thr where year(tgltrans) = '" & year(DTPicker1.Value) & "' " _
        & "AND company_code = '" & TDBCombo_company.Text & "' " & strAgama
    rs.Open SQL, CnG, adOpenDynamic, adLockReadOnly
    If rs.RecordCount > 0 Then
        vtglTrans = Format(rs!tgltrans, "yyyy-MM-dd")
        MsgBox "This Year Transactions is Already Exists on " & vtglTrans & "...!!", vbCritical, "Warning"
        rs.Close
        Me.MousePointer = vbNormal
        
        vExist = True
        Exit Sub
    End If
    rs.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    SQL = "SELECT employee_code,employee_name,a.title_code,c.title_name,a.department_code,d.department_name," _
            & "DATE(a.start_working) start_working,a.religion," _
            & "IFNULL((SELECT basic_salary FROM m_salary_standard " _
                        & "Where employee_code = a.employee_code And salary_date <= '" & Format(DateAdd("m", -1, DTPicker1), "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1),0) as basic_salary," _
            & "DATE(start_working) start_working " _
        & "FROM m_employee a JOIN m_department d ON a.department_code = d.department_code " _
        & "JOIN m_division b ON a.division_code = b.division_code AND a.company_code = b.company_code " _
        & "JOIN m_title c ON a.title_code = c.title_code " _
        & "WHERE a.company_code = '" & TDBCombo_company.Text & "' " & strAgama
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        CnG.BeginTrans
        rs.MoveFirst
        While Not rs.EOF
            
            vStartWorking = Format(rs!start_working, "yyyy-MM-dd")
            
            tglProses = Format(DTPicker1.Value, "yyyy-MM-dd")
            
            jmlBulan = DateDiff("d", vStartWorking, tglProses)
            jmlBulan = roundUp(jmlBulan / 30)
                                            
            If jmlBulan < 3 Then
                jmlTHR = 0
            ElseIf jmlBulan >= 3 And jmlBulan < 12 Then
                jmlTHR = (rs!basic_salary / 12) * jmlBulan
            Else
                jmlTHR = rs!basic_salary
            End If
            
            SQL = "INSERT INTO t_thr (company_code,tgltrans,employee_code,jmlthr,jenis,tglinput,userinput,masakerja,religion) " _
                & "VALUES " _
                & "('" & TDBCombo_company.Text & "'," _
                & "'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & rs!employee_code & "'," _
                & "'" & Round(jmlTHR) & "',0,now(),'" & LOGIN_CODE & "','" & jmlBulan & "','" & IIf(IsNull(rs!religion), 0, rs!religion) & "')"
            CnG.Execute SQL
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
    
    If cbo_religion.ListIndex = 6 Then
        strAgama = ""
    Else
        strAgama = " AND a.religion = '" & cbo_religion.ListIndex & "'"
    End If
    
'    access = IIf(LOGIN_LEVEL = 100, "", "AND (a.managerial_access = 0 OR a.managerial_access IS NULL)")
    
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    SQL = "SELECT a.nik,employee_name,a.title_code,c.title_name,a.department_code,d.department_name," _
            & "DATE(a.start_working) start_working," _
            & "d.jmlthr,d.pph21_value,(d.jmlthr - d.pph21_value) actual, " _
            & "IFNULL((SELECT basic_salary FROM m_salary_standard " _
                    & "Where employee_code = a.employee_code And salary_date <= '" & Format(DateAdd("m", -1, DTPicker1), "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1),0) as basic_salary," _
            & "CASE WHEN d.masakerja > 12 THEN 12 ELSE d.masakerja END masakerja " _
        & "FROM m_employee a JOIN m_department d ON a.department_code = d.department_code " _
        & "JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
        & "JOIN m_title c ON a.title_code = c.title_code " _
        & "LEFT JOIN (SELECT employee_code,jmlthr,masakerja,pph21_value FROM t_thr " _
            & "WHERE month(tgltrans) = '" & month(DTPicker1.Value) & "' AND year(tgltrans) = '" & year(DTPicker1.Value) & "') d on a.employee_code = d.employee_code " _
        & "WHERE a.company_code = '" & TDBCombo_company.Text & "' " & strAgama
        
'       & "IFNULL((SELECT salary FROM m_salary WHERE employee_code = a.employee_code ORDER BY salary_date DESC LIMIT 1),0) basic_salary," _

    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME & vbTab & rs!start_working _
                & vbTab & rs!DEPARTMENT_CODE & vbTab & rs!department_name _
                & vbTab & rs!title_code & vbTab & rs!title_name & vbTab & rs!basic_salary & vbTab & rs!masakerja _
                & vbTab & rs!jmlTHR & vbTab & rs!pph21_value & vbTab & rs!actual
        rs.MoveNext
        Wend
    End If
    rs.Close
    LynxGrid1.Redraw = True

End Sub

Private Sub HitungPPh(str_employee_code As String)
Dim vJmlBln As Integer
Dim vJmlBruto As Double
Dim vJmlJamsostek As Double
Dim vAvgBruto As Double
Dim vAvgJamsostek As Double
Dim vGajiPokok As Double

Dim vPTKP As String
Dim vPTKP_Value As Double
Dim vMarital As Integer
Dim vSex As Integer
Dim vNoChild As Integer
Dim vStartWorking As String
Dim vJmlBlnKerja As Integer

Dim vTHR_Value As Double
Dim vBruto_Setahun As Double
Dim vPengurang_Setahun As Double
Dim vNetto_Setahun As Double
Dim vBiayaJabatan As Double
Dim vPKP As Double

Dim vProsenPajak As Double
Dim vPPh5 As Double
Dim vPPh15 As Double
Dim vPPh25 As Double
Dim vPPh30 As Double
Dim vPPh21_THR_Setahun As Double
Dim vPPh21_Gaji_Setahun As Double
Dim vPPh21 As Double

Dim i As Integer

'On Error GoTo Err
    
    CnG.BeginTrans
    If rs.State Then rs.Close
    SQL = "SELECT SUM(CASE WHEN salary_code='SU-28' THEN 1 ELSE 0 END) jmlBln," & _
            "SUM(CASE WHEN salary_code='SU-281' THEN salary_value ELSE 0 END) jamsostek," & _
            "SUM(CASE WHEN salary_code='SU-28' THEN salary_value ELSE 0 END) bruto " & _
          "FROM h_salary " & _
          "WHERE employee_code = '" & str_employee_code & "' " & _
            "AND LEFT(MONTH,4) = '" & Format(DTPicker1.Value, "yyyy") & "' " & _
            "AND (RIGHT(MONTH,2) BETWEEN '01' " & _
                "AND '" & Format(DateAdd("m", -1, DTPicker1.Value), "MM") & "') " & _
          "GROUP BY employee_code"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vJmlBln = rs.Fields(0).Value
        vJmlJamsostek = rs.Fields(1).Value
        vJmlBruto = rs.Fields(2).Value
    Else
        vJmlBln = 1
        vJmlJamsostek = 0
        vJmlBruto = 0
    End If
    rs.Close
    
    vAvgJamsostek = vJmlJamsostek / vJmlBln
    vAvgBruto = vJmlBruto / vJmlBln
    
    SQL = "SELECT a.ptkp_type, b.marital_status, b.sex, b.no_of_children, DATE(b.start_working) " & _
          "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
          "WHERE a.employee_code = '" & str_employee_code & "' " & _
            "AND a.salary_date <= '" & Format(DateAdd("m", -1, DTPicker1.Value), "yyyy-MM-dd") & "' " & _
          "ORDER BY a.salary_date DESC LIMIT 1"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vPTKP = rs.Fields(0).Value
        vMarital = rs.Fields(1).Value
        vSex = rs.Fields(2).Value
        vNoChild = rs.Fields(3).Value
        vStartWorking = rs.Fields(4).Value
    Else
        vPTKP = "STD"
        vMarital = 0
        vSex = 0
        vNoChild = 0
        vStartWorking = Now
    End If
    rs.Close
    
    SQL = "SELECT f_get_ptkp(" & vMarital & ", " & vNoChild & "," & vSex & ", 1,'" & vPTKP & "') ptkp_value"
    rs.Open SQL, CnG, adOpenForwardOnly
    
    If rs.RecordCount > 0 Then
        vPTKP_Value = rs!ptkp_value
    End If
    rs.Close
    
    SQL = "SELECT SUM(jmlthr) FROM t_thr " & _
          "WHERE employee_code = '" & str_employee_code & "' " & _
            "AND YEAR(tgltrans) = '" & Format(DTPicker1.Value, "yyyy") & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vTHR_Value = rs.Fields(0).Value
    Else
        vTHR_Value = 0
    End If
    rs.Close
        
    If optTax(0).Value Then
        SQL = "SELECT salary_value AS basic_salary " & _
              "FROM h_salary " & _
              "WHERE employee_code = '" & str_employee_code & "' " & _
                "AND LEFT(MONTH,4) = '" & Format(DTPicker1.Value, "yyyy") & "' " & _
                "AND (RIGHT(MONTH,2) = '" & Format(DateAdd("m", -1, DTPicker1.Value), "MM") & "') " & _
                "AND salary_code = 'SU-01' " & _
              "GROUP BY employee_code"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            vGajiPokok = rs!basic_salary
        Else
            vGajiPokok = 0
        End If
        rs.Close
        
        vJmlBlnKerja = (DateDiff("m", vStartWorking, year(DTPicker1.Value) & "-12-31")) + 1
        If vJmlBlnKerja < 12 And Format(vStartWorking, "yyyy") = Format(DTPicker1.Value, "yyyy") Then
            vBruto_Setahun = IIf(i = 0, ((vAvgBruto * vJmlBlnKerja) + vTHR_Value), (vAvgBruto * vJmlBlnKerja))
            vBiayaJabatan = IIf(0.05 * vBruto_Setahun > 6000000, 6000000, 0.05 * vBruto_Setahun)
            vPengurang_Setahun = vBiayaJabatan + (vJmlBlnKerja * vAvgJamsostek)
        Else
            vBruto_Setahun = IIf(i = 0, ((vAvgBruto * 12) + vTHR_Value), (vAvgBruto * 12))
            vBiayaJabatan = IIf(0.05 * vBruto_Setahun > 6000000, 6000000, 0.05 * vBruto_Setahun)
            vPengurang_Setahun = vBiayaJabatan + (12 * vAvgJamsostek)
        End If
        
        vNetto_Setahun = vBruto_Setahun - vPengurang_Setahun
        vPKP = vNetto_Setahun - vPTKP_Value
        vPKP = Int(vPKP / 1000) * 1000
            
        If vPKP < 50000000 Then
            vProsenPajak = 0.05
            vPPh21 = 0.05 * vGajiPokok
        ElseIf vPKP > 50000000 And vPKP < 250000000 Then
            vProsenPajak = 0.15
            vPPh21 = 0.15 * vGajiPokok
        ElseIf vPKP > 250000000 And vPKP < 500000000 Then
            vProsenPajak = 0.25
            vPPh21 = 0.25 * vGajiPokok
        Else
            vProsenPajak = 0.35
            vPPh21 = 0.35 * vGajiPokok
        End If
    Else
        For i = 0 To 1
            vJmlBlnKerja = (DateDiff("m", vStartWorking, year(DTPicker1.Value) & "-12-31")) + 1
            If vJmlBlnKerja < 12 And Format(vStartWorking, "yyyy") = Format(DTPicker1.Value, "yyyy") Then
                vBruto_Setahun = IIf(i = 0, ((vAvgBruto * vJmlBlnKerja) + vTHR_Value), (vAvgBruto * vJmlBlnKerja))
                vBiayaJabatan = IIf(0.05 * vBruto_Setahun > 6000000, 6000000, 0.05 * vBruto_Setahun)
                vPengurang_Setahun = vBiayaJabatan + (vJmlBlnKerja * vAvgJamsostek)
            Else
                vBruto_Setahun = IIf(i = 0, ((vAvgBruto * 12) + vTHR_Value), (vAvgBruto * 12))
                vBiayaJabatan = IIf(0.05 * vBruto_Setahun > 6000000, 6000000, 0.05 * vBruto_Setahun)
                vPengurang_Setahun = vBiayaJabatan + (12 * vAvgJamsostek)
            End If
            
            vNetto_Setahun = vBruto_Setahun - vPengurang_Setahun
            vPKP = vNetto_Setahun - vPTKP_Value
            vPKP = Int(vPKP / 1000) * 1000
    
            If vPKP < 50000000 Then
                vProsenPajak = 0.05
                vPPh5 = 0.05 * vPKP
                vPPh15 = 0
                vPPh25 = 0
                vPPh30 = 0
            ElseIf vPKP > 50000000 And vPKP < 250000000 Then
                vProsenPajak = 0.15
                vPPh5 = 0.05 * 50000000
                vPPh15 = 0.15 * (vPKP - 50000000)
                vPPh25 = 0
                vPPh30 = 0
            ElseIf vPKP > 250000000 And vPKP < 500000000 Then
                vProsenPajak = 0.25
                vPPh5 = 0.05 * 50000000
                vPPh15 = 0.15 * 200000000
                vPPh25 = 0.25 * (vPKP - 250000000)
                vPPh30 = 0
            Else
                vProsenPajak = 0.35
                vPPh5 = 0.05 * 50000000
                vPPh15 = 0.15 * 200000000
                vPPh25 = 0.25 * 250000000
                vPPh30 = 0.35 * (vPKP - 500000000)
            End If
                
            If i = 0 Then
                vPPh21_THR_Setahun = vPPh5 + vPPh15 + vPPh25 + vPPh30
            Else
                vPPh21_Gaji_Setahun = vPPh5 + vPPh15 + vPPh25 + vPPh30
            End If
        Next
        
        vPPh21 = vPPh21_THR_Setahun - vPPh21_Gaji_Setahun
    End If
    
    '------------------------------- Update THR ------------------------------
    SQL = "UPDATE t_thr SET pph21_percentage = '" & vProsenPajak & "'," & _
            "pph21_value = '" & vPPh21 & "' " & _
          "WHERE employee_code = '" & str_employee_code & "' " & _
            "AND YEAR(tgltrans) = '" & Format(DTPicker1.Value, "yyyy") & "' " & _
            "AND MONTH(tgltrans) = '" & Format(DTPicker1.Value, "MM") & "'"
    CnG.Execute SQL
    '-------------------------------------------------------------------------
    
    CnG.CommitTrans
    Exit Sub
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Width = Me.Width - 500
    Frame1.Height = Me.Height - 3000
    LynxGrid1.Height = Frame1.Height - 400
    LynxGrid1.Width = Frame1.Width - 400
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
        
        isiGridTHR
    End If
End Sub

Private Sub timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub vbButton1_Click()
    Me.MousePointer = vbHourglass
    
    Call isiTHR
    
    If vExist = False Then
        If cbo_religion.ListIndex = 6 Then
            strAgama = ""
        Else
            strAgama = " AND religion = '" & cbo_religion.ListIndex & "'"
        End If
        
        If rscari.State Then rscari.Close
        SQL = "SELECT employee_code " & _
                 "FROM m_employee " & _
                 "WHERE company_code = '" & TDBCombo_company.Text & "' " & _
                    "AND flag_active <> 0 " & strAgama
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            rscari.MoveFirst
            While Not rscari.EOF
                Call HitungPPh(rscari.Fields(0).Value)
                rscari.MoveNext
            Wend
        End If
        rscari.Close
        
        Call isiGridTHR
    End If
    
    Me.MousePointer = vbNormal
                
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
        SQL = "DELETE FROM t_thr WHERE month(tgltrans) = '" & month(DTPicker1.Value) & "' " & _
                "AND year(tgltrans) = '" & year(DTPicker1.Value) & "' " & strAgama & _
                "AND company_code = '" & TDBCombo_company.Text & "'"
                'AND kddivisi = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "'"
        CnG.Execute SQL
        isiGridTHR
    End If
    Me.MousePointer = vbNormal
End Sub

Private Sub vbButton3_Click()
    isiGridTHR
End Sub
