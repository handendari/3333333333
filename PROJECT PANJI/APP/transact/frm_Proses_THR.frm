VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_proses_thr 
   Caption         =   "PROSES THR"
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
   Begin VB.TextBox txt_division_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   1950
      Width           =   3855
   End
   Begin VB.ComboBox cbo_religion 
      Height          =   315
      ItemData        =   "frm_Proses_THR.frx":058A
      Left            =   8520
      List            =   "frm_Proses_THR.frx":05A3
      TabIndex        =   6
      Top             =   1950
      Width           =   2265
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1410
      OleObjectBlob   =   "frm_Proses_THR.frx":05EC
      TabIndex        =   2
      Top             =   1590
      Width           =   1785
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1590
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Height          =   8985
      Left            =   120
      TabIndex        =   0
      Top             =   2370
      Width           =   20025
      Begin prj_panji.LynxGrid LynxGrid1 
         Height          =   8325
         Left            =   60
         TabIndex        =   8
         Top             =   150
         Width           =   14685
         _ExtentX        =   25903
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
      Left            =   1410
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   160366595
      UpDown          =   -1  'True
      CurrentDate     =   39278
   End
   Begin prj_panji.vbButton vbButton3 
      Height          =   495
      Left            =   3210
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
      MICON           =   "frm_Proses_THR.frx":25AA
      PICN            =   "frm_Proses_THR.frx":25C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.vbButton vbButton2 
      Height          =   465
      Left            =   12750
      TabIndex        =   10
      Top             =   1800
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&Hapus"
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
      MICON           =   "frm_Proses_THR.frx":3658
      PICN            =   "frm_Proses_THR.frx":3674
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.vbButton vbButton1 
      Height          =   465
      Left            =   11310
      TabIndex        =   11
      Top             =   1800
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&Proses"
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
      MICON           =   "frm_Proses_THR.frx":4706
      PICN            =   "frm_Proses_THR.frx":4722
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_division 
      Height          =   375
      Left            =   1410
      OleObjectBlob   =   "frm_Proses_THR.frx":57B4
      TabIndex        =   14
      Top             =   1950
      Width           =   1785
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DIVISI"
      Height          =   195
      Left            =   900
      TabIndex        =   15
      Top             =   1980
      Width           =   465
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PROSES THR"
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
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "AGAMA"
      Height          =   195
      Left            =   7620
      TabIndex        =   7
      Top             =   2010
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TANGGAL"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   1140
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   255
      TabIndex        =   4
      Top             =   1650
      Width           =   1110
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_Proses_THR.frx":7773
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16860
   End
End
Attribute VB_Name = "frm_proses_thr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rsDiv As New ADODB.Recordset

Dim strsql As String
Dim access As String

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
    
    Call createGrid
    
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
      .AddColumn "KODE KARY.", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "NAMA KARY.", 3000, , , , , , , , True
      .AddColumn "TGL MASUK KERJA", 2000, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , True
      .AddColumn "Title Code", 1000, , , , , , , , False
      .AddColumn "JABATAN", 2000, , , , , , , , , True
      '.AddColumn "Basic Income", 1500, lgAlignRightCenter, lgNumeric, "#,##", , , , , True
      .AddColumn "NILAI THR", 1500, lgAlignRightCenter, lgNumeric, "#,##", , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiCompany()
    Dim rsCompany As New ADODB.Recordset
    
    strsql = "select * from m_company order by company_code"
    rsCompany.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo_company.RowSource = rsCompany
    'rs.Close

End Sub

Public Sub load_data_division()
    TDBCombo_division.Text = "": txt_division_name.Text = ""
    
    If rsDiv.State Then rsDiv.Close
    SQL = "select * from m_division where company_code = '" & TDBCombo_company.Text & "' order by company_code"
    rsDiv.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_division.RowSource = rsDiv
End Sub

Public Sub isiTHR()
Dim rs2 As New ADODB.Recordset

Dim i As Integer
Dim jmlTHR As Double
Dim strAgama As String
Dim tglProses As String
Dim jmlBulan As Double
Dim blnTHR As String

Dim vEmployee_Code As String
Dim vUMK As Double
Dim vThrCode As String
Dim vTHR_HL As Double
Dim vFlag_percentage As Integer
Dim vTHR_Naik As Double
Dim vTipe_Thr As String
Dim vTHR_Upper As Double
Dim vTHR_Under As Double
Dim vFlag_proporsional As Integer
Dim vFlag_UMK As Integer
Dim vBasic_Salary As Double
Dim vLast_THR As Double

    Me.MousePointer = vbHourglass
    
    SQL = "SELECT * FROM m_pref_umk WHERE date(umk_date) <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
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
    strsql = "SELECT 1 FROM t_thr where month(tgltrans) = '" & month(DTPicker1.Value) & "' AND year(tgltrans) = '" & year(DTPicker1.Value) & "' " _
        & "AND " & IIf(TDBCombo_division.Text = "", "kdcabang = '" & TDBCombo_company.Text & "'", _
                    "kdcabang = '" & TDBCombo_company.Text & "' AND kddivisi = '" & TDBCombo_division.Text & "'") & " " _
        & strAgama
    rs.Open strsql, CnG, adOpenDynamic, adLockReadOnly
    If rs.RecordCount > 0 Then
        MsgBox "Transaksi Tahun Ini Sudah Terjadi...", vbCritical, "Warning"
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
            & "DATE(start_working) start_working, " _
            & "IFNULL((SELECT thr_code FROM m_thr_status f WHERE f.status_code = a.status_code),'') thr_code " _
        & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code AND a.company_code = b.company_code " _
        & "JOIN m_title c ON a.title_code = c.title_code " _
        & "WHERE " & IIf(TDBCombo_division.Text = "", "a.company_code = '" & TDBCombo_company.Text & "'", _
                    "a.company_code = '" & TDBCombo_company.Text & "' AND a.division_code = '" & TDBCombo_division.Text & "'") & " " & strAgama
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        CnG.BeginTrans
        rs.MoveFirst
        While Not rs.EOF
            blnTHR = DateAdd("m", 0, DTPicker1.Value)
            
'            If month(DTPicker1) < 10 Then
'                tglProses = year(blnTHR) & "-0" & month(blnTHR) & "-" & getEndDay(month(blnTHR), year(blnTHR))
'            Else
'                tglProses = year(blnTHR) & "-" & month(blnTHR) & "-" & getEndDay(month(blnTHR), year(blnTHR))
'            End If
            
            tglProses = DTPicker1.Value
            vEmployee_Code = rs!employee_code
            
            jmlBulan = DateDiff("M", rs!start_working, tglProses)
            
            vThrCode = rs!thr_code
            
            If vThrCode <> "" Then
                SQL = "SELECT thr_value, flag_percentage, thr_naik, " & _
                        "thr_under, thr_upper, flag_proporsional, flag_umk FROM m_thr_detail " & _
                      "WHERE thr_under < '" & jmlBulan & "' " & _
                        "AND thr_upper >= '" & jmlBulan & "' " & _
                        "AND thr_code = '" & vThrCode & "'"
                rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
                If rs2.RecordCount > 0 Then
                    vTHR_HL = rs2!thr_value
                    vFlag_percentage = rs2!flag_percentage
                    vTHR_Naik = rs2!thr_naik
                    vTHR_Under = rs2!thr_under
                    vTHR_Upper = rs2!thr_upper
                    vFlag_proporsional = IIf(IsNull(rs2!flag_proporsional), 0, rs2!flag_proporsional)
                    vFlag_UMK = IIf(IsNull(rs2!flag_umk), 0, rs2!flag_umk)
                End If
                rs2.Close
                
                SQL = "SELECT IFNULL((SELECT basic_salary FROM m_salary_standard WHERE employee_code = '" & vEmployee_Code & "' ORDER BY salary_date DESC LIMIT 1),0) basic_salary"
                rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    vBasic_Salary = rs2!basic_salary
                End If
                rs2.Close
                
                If vFlag_proporsional <> 0 Then
                    If vFlag_UMK = 0 Then
                        If vFlag_percentage <> 0 Then
                            jmlTHR = (jmlBulan / 12) * vUMK
                        Else
                            jmlTHR = vUMK
                        End If
                    Else
                        If vFlag_percentage <> 0 Then
                            jmlTHR = (jmlBulan / 12) * vBasic_Salary
                        Else
                            jmlTHR = vBasic_Salary
                        End If
                    End If
                Else
                    SQL = "SELECT COUNT(*) jml FROM t_thr " & _
                          "WHERE employee_code = '" & vEmployee_Code & "' " & _
                            "AND bln_kerja > '" & vTHR_Under & "' AND bln_kerja <= '" & vTHR_Upper & "'"
                    rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs2.RecordCount > 0 Then
                        i = rs2!jml
                    Else
                        i = 0
                    End If
                    rs2.Close
                    
                    If i = 0 Then
                        If vFlag_UMK = 0 Then
                            If vFlag_percentage <> 0 Then
                                jmlTHR = (vTHR_HL / 100) * vUMK
                            Else
                                jmlTHR = vTHR_HL
                            End If
                        Else
                            If vFlag_percentage <> 0 Then
                                jmlTHR = (vTHR_HL / 12) * vBasic_Salary
                            Else
                                jmlTHR = vTHR_HL
                            End If
                        End If
                    Else
                        SQL = "SELECT jmlthr FROM t_thr " & _
                              "WHERE employee_code = '" & vEmployee_Code & "' " & _
                                "AND bln_kerja > '" & vTHR_Under & "' AND bln_kerja <= '" & vTHR_Upper & "' " & _
                              "ORDER BY tgltrans DESC LIMIT 1"
                        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                        
                        If rs2.RecordCount > 0 Then
                            vLast_THR = rs2!jmlTHR
                        End If
                        rs2.Close
                        
                        If vFlag_UMK = 0 Then
                            If vFlag_percentage <> 0 Then
                                jmlTHR = ((vTHR_HL / 100) * vUMK) + vTHR_Naik
                            Else
                                jmlTHR = vLast_THR + vTHR_Naik
                            End If
                        Else
                            If vFlag_percentage <> 0 Then
                                jmlTHR = ((vTHR_HL / 12) * vBasic_Salary) + vTHR_Naik
                            Else
                                jmlTHR = vLast_THR + vTHR_Naik
                            End If
                        End If
                    End If
                End If
            End If
            
            If TDBCombo_division = "" Then
                Dim vDivCode As String
                
                SQL = "SELECT division_code FROM m_employee WHERE employee_code = '" & vEmployee_Code & "'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vDivCode = rscari!division_code
                End If
                rscari.Close
                
                strsql = "INSERT INTO t_thr (kdcabang,kddivisi,tgltrans,employee_code,jmlthr,bln_kerja,jenis,tglinput,userinput,masakerja,religion) " _
                    & "VALUES " _
                    & "('" & TDBCombo_company.Text & "','" & vDivCode & "'," _
                    & "'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & rs!employee_code & "'," _
                    & "'" & jmlTHR & "','" & jmlBulan & "',0,now(),'" & LOGIN_CODE & "','" & jmlBulan & "','" & IIf(IsNull(rs!religion), 0, rs!religion) & "')"
            Else
                strsql = "INSERT INTO t_thr (kdcabang,kddivisi,tgltrans,employee_code,jmlthr,bln_kerja,jenis,tglinput,userinput,masakerja,religion) " _
                    & "VALUES " _
                    & "('" & TDBCombo_company.Text & "','" & TDBCombo_division.Text & "'," _
                    & "'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & rs!employee_code & "'," _
                    & "'" & jmlTHR & "','" & jmlBulan & "',0,now(),'" & LOGIN_CODE & "','" & jmlBulan & "','" & IIf(IsNull(rs!religion), 0, rs!religion) & "')"
            End If
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
        MsgBox "Agama Tidak Valid...", vbExclamation, "Error"
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
    strsql = "SELECT a.employee_code,employee_name,a.title_code,c.title_name," _
            & "DATE(a.start_working) start_working," _
            & "d.jmlthr,d.masakerja " _
        & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
        & "JOIN m_title c ON a.title_code = c.title_code " _
        & "LEFT JOIN (SELECT employee_code,jmlthr,masakerja FROM t_thr " _
            & "WHERE month(tgltrans) = '" & month(DTPicker1.Value) & "' AND year(tgltrans) = '" & year(DTPicker1.Value) & "') d on a.employee_code = d.employee_code " _
        & "WHERE " & IIf(TDBCombo_division.Text = "", "a.company_code = '" & TDBCombo_company.Text & "'", _
                    "a.company_code = '" & TDBCombo_company.Text & "' AND a.division_code = '" & TDBCombo_division.Text & "'") & " " & access & " AND a.company_code = '" & TDBCombo_company.Text & "' " & strAgama
        
'       & "IFNULL((SELECT salary FROM m_salary WHERE employee_code = a.employee_code ORDER BY salary_date DESC LIMIT 1),0) basic_salary," _

    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!employee_code & vbTab & rs!EMPLOYEE_NAME & vbTab & rs!start_working _
                & vbTab & rs!title_code & vbTab & rs!title_name & vbTab & rs!jmlTHR
        rs.MoveNext
        Wend
    End If
    rs.Close
    LynxGrid1.Redraw = True

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


Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
        
        Call load_data_division
    End If
End Sub

Private Sub TDBCombo_division_ItemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
        txt_division_name.Text = TDBCombo_division.Columns("division_name").Value
        
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
    
    tanya = MsgBox("Data Tidak Akan Kembali.." & Chr(13) & "Apakah Yakin Akan Menghapus Data??", vbExclamation + vbYesNo, "Warning")
    If tanya = vbYes Then
        strsql = "DELETE FROM t_thr WHERE month(tgltrans) = '" & month(DTPicker1.Value) & "' " & _
                "AND year(tgltrans) = '" & year(DTPicker1.Value) & "' " & strAgama & _
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
