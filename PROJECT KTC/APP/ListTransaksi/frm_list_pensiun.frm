VERSION 5.00
Begin VB.Form frm_List_pensiun 
   Caption         =   "List Pension"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin prj_fej_jkt.vbButton vbButton1 
      Height          =   345
      Left            =   4950
      TabIndex        =   2
      Top             =   120
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Refresh Data"
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
      MICON           =   "frm_list_pensiun.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   9825
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   20025
      Begin prj_fej_jkt.LynxGrid LynxGrid1 
         Height          =   8835
         Left            =   150
         TabIndex        =   1
         Top             =   210
         Width           =   15765
         _ExtentX        =   27808
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
   Begin prj_fej_jkt.vbButton vbButton2 
      Height          =   345
      Left            =   630
      TabIndex        =   3
      Top             =   120
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Add New"
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
      MICON           =   "frm_list_pensiun.frx":001C
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
      Left            =   2070
      TabIndex        =   4
      Top             =   120
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Edit Data"
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
      MICON           =   "frm_list_pensiun.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_fej_jkt.vbButton vbButton4 
      Height          =   345
      Left            =   3510
      TabIndex        =   5
      Top             =   120
      Width           =   1425
      _ExtentX        =   2514
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
      MICON           =   "frm_list_pensiun.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_fej_jkt.vbButton vbButton5 
      Height          =   345
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Print Report"
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
      MICON           =   "frm_list_pensiun.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_fej_jkt.vbButton vbButton6 
      Height          =   345
      Left            =   8310
      TabIndex        =   7
      Top             =   120
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Print PPh 21 Pensiun"
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
      MICON           =   "frm_list_pensiun.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_fej_jkt.vbButton vbButton7 
      Height          =   345
      Left            =   10080
      TabIndex        =   8
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Print PPh 21 Tahunan"
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
      MICON           =   "frm_list_pensiun.frx":00A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frm_List_pensiun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim strsql As String

Private Sub Form_Load()
    createGrid
    isiGridAbsen
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "Date", 1500, lgAlignCenterCenter, lgDate, "yyyy-MM-dd", , , , , True
      .AddColumn "NIK", 2000, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Employee Name", 3000, , , , , , , , True
      .AddColumn "Div. Code", , , , , , , , , False
      .AddColumn "Division", 2500, , , , , , , , , True
      .AddColumn "Title Code", , , , , , , , , False
      .AddColumn "Title", 2500, , , , , , , , , True
      '.AddColumn "Value", 2200, lgAlignCenterCenter, lgNumeric, "#,##", , , , , True
      .AddColumn "Remark", 2300, , , , , , , , True
      .AddColumn "Entry Date", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Public Sub isiGridAbsen()
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    'If rs.State = 1 Then rs.Close
    strsql = "select a.tgltrans,a.employee_code,b.employee_name," & _
                "b.division_code,c.division_name,b.title_code,d.title_name,a.ket,a.tglinput " & _
            "from t_pensiun a join m_employee b on a.employee_code = b.employee_code " & _
            "left join m_division c on b.division_code = c.division_code AND b.company_code = c.company_code " & _
            "join m_title d on b.title_code = d.title_code order by a.tgltrans, a.employee_code"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!tgltrans & vbTab & rs!employee_code & vbTab & rs!employee_name _
                & vbTab & rs!division_code & vbTab & rs!division_name _
                & vbTab & rs!title_code & vbTab & rs!title_name _
                 & vbTab & rs!ket & vbTab & rs!tglinput
            rs.MoveNext
        Wend
    End If
    
    vbButton1.Enabled = IIf(rs.RecordCount = 0, False, True)
    vbButton3.Enabled = IIf(rs.RecordCount = 0, False, True)
    vbButton4.Enabled = IIf(rs.RecordCount = 0, False, True)
    vbButton5.Enabled = IIf(rs.RecordCount = 0, False, True)
    vbButton6.Enabled = IIf(rs.RecordCount = 0, False, True)
    vbButton7.Enabled = IIf(rs.RecordCount = 0, False, True)
    
    rs.Close
    LynxGrid1.Redraw = True
End Sub

Private Sub showData()
Dim rs2 As New ADODB.Recordset
'    If LynxGrid1.Rows > 0 Then
'        With frm_trans_pensiun
''            .DTPicker1.Value = LynxGrid1.CellValue(LynxGrid1.Row, 0)
'            .txtkdkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
'            .txtnmkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 2)
''            .txtkddiv.Text = LynxGrid1.CellText(LynxGrid1.Row, 3)
''            .txtdivision.Text = LynxGrid1.CellText(LynxGrid1.Row, 4)
''            .txtkdtitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 5)
''            .txttitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 6)
''            .txt_jml_pensiun.Value = LynxGrid1.CellValue(LynxGrid1.Row, 7)
''            .txtket.Text = LynxGrid1.CellValue(LynxGrid1.Row, 8)
'            .editTrans = True
'            .Show vbModal
'        End With
'    End If
strsql = "SELECT a.tgltrans," _
                & "a.employee_code,b.employee_name," _
                & "b.company_code,c.company_name," _
                & "b.department_code,d.department_name," _
                & "b.division_code,e.division_name," _
                & "b.title_code,f.title_name,b.start_working," _
                & "a.basic_salary,a.kali_gaji,a.uang_pensiun," _
                & "a.persen_transport,a.lain1,a.lain2,a.ket " _
        & "FROM t_pensiun a JOIN m_employee b ON a.employee_code = b.employee_code " _
        & "JOIN m_company c ON b.company_code = c.company_code " _
        & "JOIN m_department d ON b.department_code = d.department_code and b.company_code = c.company_code " _
        & "JOIN m_division e ON b.division_code = e.division_code and b.department_code = d.department_code and b.company_code = c.company_code " _
        & "JOIN m_title f ON b.title_code = f.title_code  " _
        & "WHERE a.employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 1) & "'"
rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs2.RecordCount > 0 Then
    If LynxGrid1.Rows > 0 Then
        With frm_trans_pensiun
            .DTPicker1.Enabled = False
            .DTPicker1.Value = rs2!tgltrans
            .txtkdkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
            .txtnmkar.Text = rs2!employee_name
            .txtsite_code.Text = rs2!COMPANY_CODE
            .txt_company_name.Text = rs2!company_name
            .txtdept_code.Text = rs2!department_code
            .txtnmDept.Text = rs2!department_name
            .txtkddiv.Text = rs2!division_code
            .txtdivision.Text = rs2!division_name
            .txtkdtitle.Text = rs2!title_code
            .txttitle.Text = rs2!title_name
            .txtstart_working.Text = rs2!start_working
            .txtgaji_bersih.Value = rs2!basic_salary
            .txt_kali.Value = rs2!kali_gaji
            .txt_jml_pesangon.Value = rs2!uang_pensiun
            .txt_15_persen.Value = rs2!persen_transport
            .txt_other1.Value = rs2!lain1
            .txt_other2.Value = rs2!lain2
            .txtket.Text = rs2!ket
            .editTrans = True
            .Show vbModal
        End With
    End If
End If
rs2.Close
End Sub

Private Sub newData()
    With frm_trans_pensiun
        .editTrans = False
        .Show vbModal
    End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Width = Me.Width - 500
    Frame1.Height = Me.Height - 1200
    LynxGrid1.Height = Frame1.Height - 400
    LynxGrid1.Width = Frame1.Width - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_List_pensiun = Nothing
End Sub

Private Sub LynxGrid1_DblClick()
    showData
End Sub

Private Sub LynxGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then showData
End Sub

Private Sub vbButton1_Click()
    isiGridAbsen
End Sub

Private Sub vbButton2_Click()
    newData
End Sub


Private Sub vbButton3_Click()
    showData
End Sub

Private Sub vbButton4_Click()
    Dim tanya As Integer
    tanya = MsgBox("Deleted Data Can Not be Undo, Are You Sure?", vbCritical + vbYesNo, "Warning")
    If tanya = vbYes Then
        CnG.Execute "DELETE FROM t_pensiun where employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 1) & "'"
        CnG.Execute "UPDATE m_employee SET flag_active = 1 WHERE employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 1) & "'"
        CnG.Execute "DELETE FROM h_salary_new where employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 1) & "' " _
                & "AND month = '" & Mid(Format(LynxGrid1.CellText(LynxGrid1.Row, 0), "yyyy-MM-dd"), 1, 7) & "'"
        MsgBox "Delete Data Sucsess..."
        isiGridAbsen
    End If
End Sub

Private Sub vbButton5_Click()
'    frm_rpt_compensation.Show
Dim strsql As String
Dim pTgl1 As String, pTgl2 As String, str_file As String, kdkaryawan As String
Dim a As New frm_rpt
Dim judul As String
Dim str_param_periode As String
    
'    If check_validate_tdbcombo(TDBCombo_company) = False Then
'        MsgBox "No Branch Office selected!", vbInformation, headerMSG
'        Exit Sub
'    End If

'    If SSTab1.Tab = 0 Then
'        pTgl1 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-01"
'        pTgl2 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-" & getEndDay(DTPicker_monthly.Month)
'        kdkaryawan = txt_monthly_employee_code.Text
'        judul = txt_title.Text
'        str_param_periode = "MONTH : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
'    Else
'        pTgl1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
'        pTgl2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
'        kdkaryawan = txt_periode_employee_code.Text
'        judul = txt_title1.Text
'        str_param_periode = "PERIODE : " & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & " s/d " & "PERIODE : " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
'    End If
    
    str_file = "\report\rpt_kompensasi.rpt"
    
    strsql = "call sp_Kompensasi('" & LynxGrid1.CellText(LynxGrid1.Row, 1) & "','" & UCase(LynxGrid1.CellText(LynxGrid1.Row, 7)) & "')"
    Call a.Show
    a.Caption = "REPORT COMPENSATION"
    'str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
     
    Call a.rpt_view(strsql, str_file, str_param_periode)
End Sub

Private Sub vbButton6_Click()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2 As String

'If check_validate_tdbcombo(TDBCombo_company) = False Then
'    MsgBox "No Company selected!", vbInformation, headerMSG
'    Exit Sub
'End If

str_file = "\report\rpt_spt_pph21_pesangon.rpt"
str_employee_code = LynxGrid1.CellText(LynxGrid1.Row, 1)


d1 = Format(LynxGrid1.CellText(LynxGrid1.Row, 0), "yyyy")
'd2 = Format(DTPicker_yearly.Value, "yyyy-12-27")

'str_sql = "CALL spg_spt21_ebiz('" & str_employee_code & "', '" & d1 & "', '" & d2 & "', '" & d2 & "');"
'CnG.Execute str_sql
str_sql = "CALL spr_pph21_pesangon('" & str_employee_code & "','" & d1 & "');"
str_param_periode = "YEARLY : (" & d1 & ")"

Call a.Show
a.Caption = "SPT PPH PASAL 21 REPORT"
Call a.rpt_view_spt_pph21(str_sql, str_file, str_param_periode)

End Sub

Private Sub vbButton7_Click()
Dim str_sql, str_param_periode, str_file, strsql As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2 As String
Dim rs As New ADODB.Recordset
Dim v_company As String

str_file = "\report\rpt_spt_pph21_ebiz.rpt"
str_employee_code = LynxGrid1.CellText(LynxGrid1.Row, 1)

d1 = Format(LynxGrid1.CellText(LynxGrid1.Row, 0), "yyyy")
'd2 = Format(DTPicker_yearly.Value, "yyyy-12-27")

'str_sql = "CALL spg_spt21_ebiz('" & str_employee_code & "', '" & d1 & "', '" & d2 & "', '" & d2 & "');"
'CnG.Execute str_sql

strsql = "Select company_code FROM m_employee WHERE employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 1) & "'"
rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    v_company = rs!COMPANY_CODE
rs.Close

str_sql = "CALL spg_spt21_ebiz1('" & str_employee_code & "','" & d1 & "', " & _
        "'" & v_company & "', " & 1 & ");"
str_param_periode = "YEARLY : (" & Format(LynxGrid1.CellText(LynxGrid1.Row, 0), "yyyy") & ")"

Call a.Show
a.Caption = "SPT PPH PASAL 21 REPORT"
Call a.rpt_view_spt_pph21(str_sql, str_file, str_param_periode)
End Sub
