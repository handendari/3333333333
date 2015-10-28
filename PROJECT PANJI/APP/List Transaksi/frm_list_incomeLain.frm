VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_List_IncomeLain 
   Caption         =   "PENDAPATAN / PENGELUARAN LAIN LAIN"
   ClientHeight    =   10140
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10140
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
      Left            =   2700
      Locked          =   -1  'True
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   1740
      Width           =   3495
   End
   Begin VB.Frame fra_status_emp 
      Caption         =   "Daftar Pendapatan / Pengeluaran"
      Height          =   585
      Left            =   8850
      TabIndex        =   16
      Top             =   720
      Width           =   3315
      Begin VB.OptionButton optIncome 
         Caption         =   "Pendapatan"
         Height          =   225
         Left            =   240
         TabIndex        =   4
         Top             =   270
         Width           =   1365
      End
      Begin VB.OptionButton optExpense 
         Caption         =   "Pengeluaran"
         Height          =   225
         Left            =   1620
         TabIndex        =   5
         Top             =   270
         Width           =   1245
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   930
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   102629379
      CurrentDate     =   40809
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1440
      OleObjectBlob   =   "frm_list_incomeLain.frx":0000
      TabIndex        =   3
      Top             =   1380
      Width           =   1215
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   2700
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1380
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Height          =   7605
      Left            =   120
      TabIndex        =   0
      Top             =   2190
      Width           =   20025
      Begin prj_panji.LynxGrid LynxGrid1 
         Height          =   7515
         Left            =   120
         TabIndex        =   13
         Top             =   210
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   13256
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   3150
      TabIndex        =   2
      Top             =   930
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   160301059
      CurrentDate     =   40809
   End
   Begin prj_panji.vbButton cmdDelete 
      Height          =   465
      Left            =   10830
      TabIndex        =   8
      Top             =   1590
      Width           =   1305
      _ExtentX        =   2302
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
      MICON           =   "frm_list_incomeLain.frx":1FBE
      PICN            =   "frm_list_incomeLain.frx":1FDA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.vbButton cmdEdit 
      Height          =   465
      Left            =   9480
      TabIndex        =   7
      Top             =   1590
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&Ubah"
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
      MICON           =   "frm_list_incomeLain.frx":306C
      PICN            =   "frm_list_incomeLain.frx":3088
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.vbButton cmdNew 
      Height          =   465
      Left            =   8130
      TabIndex        =   6
      Top             =   1590
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&Tambah"
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
      MICON           =   "frm_list_incomeLain.frx":411A
      PICN            =   "frm_list_incomeLain.frx":4136
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.vbButton cmdRefresh 
      Height          =   495
      Left            =   4920
      TabIndex        =   15
      Top             =   840
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
      MICON           =   "frm_list_incomeLain.frx":51C8
      PICN            =   "frm_list_incomeLain.frx":51E4
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
      Left            =   1440
      OleObjectBlob   =   "frm_list_incomeLain.frx":6276
      TabIndex        =   18
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DIVISI :"
      Height          =   195
      Left            =   840
      TabIndex        =   19
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PENDAPATAN / PENGELUARAN LAIN LAIN"
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
      Left            =   360
      TabIndex        =   14
      Top             =   150
      Width           =   7725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "S/D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   12
      Top             =   990
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PERIODE :"
      Height          =   195
      Left            =   630
      TabIndex        =   11
      Top             =   990
      Width           =   810
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN :"
      Height          =   195
      Left            =   210
      TabIndex        =   10
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_list_incomeLain.frx":8235
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21810
   End
End
Attribute VB_Name = "frm_List_IncomeLain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsDiv As New ADODB.Recordset

'Private Sub btnImport_Click()
'frm_import_other_income.Show
'End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    LynxGrid1.Clear
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    
    createGrid
    
    Call isiCompany
    
    cmdNew.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "TANGGAL", 1200, lgAlignCenterCenter, , , , , , , True
      .AddColumn "KODE KARY.", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "NAMA KARY.", 2500, , , , , , , , True
      .AddColumn "TITLE CODE", , , , , , , , , False
      .AddColumn "JABATAN", 2000, , , , , , , , , True
      .AddColumn "TOTAL", 1200, lgAlignCenterCenter, lgNumeric, "#,##", , , , , True
      .AddColumn "TIPE", 2000, , , , , , , , True
      .AddColumn "DESKRIPSI", 2000, , , , , , , , True
      .AddColumn "TGL. INPUT", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , True
      .AddColumn "Employee Code", 2500, , , , , , , , False
      .AddColumn "Flag Income Expense", 2500, , , , , , , , False
      .AddColumn "Nomer", 2500, , , , , , , , False
      .AddColumn "Flag Type", 2500, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiCompany()
    Dim rsCompany As New ADODB.Recordset
    
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
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

Public Sub isiGridAbsen()
Dim access As String

    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    If optIncome.Value = True Then
        SQL = "SELECT a.tgltrans,b.nik,b.employee_name," _
                    & "b.title_code,d.title_name," _
                    & "a.jmlPotong,a.nm_type,a.remark, a.tglinput, b.employee_code, " _
                    & "a.flag_income_expense, a.nomer,a.flag_type " _
            & "FROM t_income_expense a JOIN m_employee b ON a.employee_code = b.employee_code " _
            & "JOIN m_division c ON b.division_code = c.division_code and b.company_code = c.company_code " _
            & "JOIN m_title d ON b.title_code = d.title_code " _
            & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                    "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
            & "AND a.tgltrans BETWEEN '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker2.Value, "yyyy-MM-dd") & "' " _
            & "AND a.flag_income_expense = 0"
    Else
        SQL = "SELECT a.tgltrans,b.nik,b.employee_name," _
                    & "b.title_code,d.title_name," _
                    & "a.jmlPotong,'OTHER EXPENSE' nm_type,a.remark, a.tglinput, b.employee_code, " _
                    & "a.flag_income_expense, a.nomer,1 flag_type " _
            & "FROM t_income_expense a JOIN m_employee b ON a.employee_code = b.employee_code " _
            & "JOIN m_division c ON b.division_code = c.division_code and b.company_code = c.company_code " _
            & "JOIN m_title d ON b.title_code = d.title_code " _
            & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                    "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
            & "AND a.tgltrans BETWEEN '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker2.Value, "yyyy-MM-dd") & "' " _
            & "AND a.flag_income_expense = 1"
    End If
    
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!tgltrans & vbTab & rs!nik & vbTab & rs!EMPLOYEE_NAME _
                & vbTab & rs!title_code & vbTab & rs!title_name _
                & vbTab & rs!jmlpotong & vbTab & rs!nm_type & vbTab & rs!remark _
                & vbTab & rs!tglinput & vbTab & rs!employee_code _
                & vbTab & rs!flag_income_expense & vbTab & rs!nomer & vbTab & rs!flag_type
            rs.MoveNext
        Wend
    End If
    
    cmdNew.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)
    cmdEdit.Enabled = IIf(rs.RecordCount = 0, False, True)
    cmdDelete.Enabled = IIf(rs.RecordCount = 0, False, True)
    cmdRefresh.Enabled = IIf(rs.RecordCount = 0, False, True)
    
    rs.Close
    LynxGrid1.Redraw = True
End Sub

Private Sub showData()
    If LynxGrid1.Rows > 0 Then
        With frm_trans_IncomeLain
            .TDBCombo_company.Text = TDBCombo_company.Text
            .txt_company_name.Text = txt_company_name.Text
            
            If TDBCombo_division.Text <> "" Then
                .txtkddiv.Text = TDBCombo_division.Text
                .txtdivision.Text = txt_division_name.Text
            End If
            
            .DTPicker1.Value = LynxGrid1.CellText(LynxGrid1.Row, 0)
            .txtkdkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 9)
            .txtnik.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
            .txtnmkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 2)
            .txtkdtitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 3)
            .txttitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 4)
            .txtjumlah.Text = FormatNumber(DropAllComma(LynxGrid1.CellValue(LynxGrid1.Row, 5)))
            
            .v_value = .txtjumlah.Text
            .v_nomer = LynxGrid1.CellText(LynxGrid1.Row, 11)
            .Combo1.ListIndex = LynxGrid1.CellText(LynxGrid1.Row, 10)
            
            If .Combo1.ListIndex = 0 Then
                .Combo2.Text = LynxGrid1.CellText(LynxGrid1.Row, 6)
                .Combo2.ListIndex = LynxGrid1.CellText(LynxGrid1.Row, 12)
            End If
            
            .txtket.Text = LynxGrid1.CellText(LynxGrid1.Row, 7)
            .editTrans = True
            .Show vbModal
        End With
    End If
End Sub

Private Sub newData()
    With frm_trans_IncomeLain
        .DTPicker1.Value = Format(DTPicker1.Value, "yyyy-MM-dd")
        .TDBCombo_company.Text = TDBCombo_company.Text
        .txt_company_name.Text = txt_company_name.Text
        
        If TDBCombo_division.Text <> "" Then
            .txtkddiv.Text = TDBCombo_division.Text
            .txtdivision.Text = txt_division_name.Text
        End If
        
        .Combo1.ListIndex = IIf(optIncome.Value, 0, 1)
        If optIncome.Value Then
            .Combo2.Enabled = True
            .Combo2.Visible = True
            .Combo2.ListIndex = 0
        Else
            .Combo2.Visible = False
        End If
        .Show vbModal
    End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Width = Me.Width - 500
    Frame1.Height = Me.Height - 3000
    LynxGrid1.Height = Frame1.Height - 400
    LynxGrid1.Width = Frame1.Width - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_List_IncomeLain = Nothing
End Sub

Private Sub LynxGrid1_DblClick()
    showData
End Sub

Private Sub LynxGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        showData
    End If
End Sub

Private Sub LynxGrid2_RowColChanged()
    isiGridAbsen
End Sub

Private Sub optExpense_Click()
    isiGridAbsen
End Sub

Private Sub optIncome_Click()
    isiGridAbsen
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
        
        optIncome.Value = True
        Call load_data_division
        Call isiGridAbsen
    End If
End Sub

Private Sub TDBCombo_division_ItemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
        txt_division_name.Text = TDBCombo_division.Columns("division_name").Value
        
        Call isiGridAbsen
    End If
End Sub

Private Sub cmdRefresh_Click()
    isiGridAbsen
End Sub

Private Sub cmdNew_Click()
    newData
End Sub


Private Sub cmdEdit_Click()
    showData
End Sub

Private Sub cmdDelete_Click()
    Dim tanya As Integer
    tanya = MsgBox("Data Akan Hilang, Apakah Yakin Ingin Menghapus Data Ini?", vbCritical + vbYesNo, "Warning")
    If tanya = vbYes Then
        CnG.Execute "DELETE FROM t_income_expense where employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 9) & "' " _
            & "and nomer = '" & LynxGrid1.CellText(LynxGrid1.Row, 11) & "' " _
            & "AND flag_income_expense = '" & IIf(optIncome.Value, 0, 1) & "'"
        MsgBox "Hapus Data Berhasil..."
        
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute SQL
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        isiGridAbsen
    End If
End Sub
