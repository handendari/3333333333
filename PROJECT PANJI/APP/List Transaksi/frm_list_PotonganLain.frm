VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_List_PotonganLain 
   Caption         =   "Others Expense History"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin prj_panji.vbButton vbButton1 
      Height          =   345
      Left            =   12600
      TabIndex        =   13
      Top             =   1410
      Width           =   1365
      _ExtentX        =   2408
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
      MICON           =   "frm_list_PotonganLain.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.vbButton vbButton4 
      Height          =   345
      Left            =   10950
      TabIndex        =   12
      Top             =   1410
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frm_list_PotonganLain.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.vbButton vbButton3 
      Height          =   345
      Left            =   9510
      TabIndex        =   11
      Top             =   1410
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
      MICON           =   "frm_list_PotonganLain.frx":0038
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
      Height          =   345
      Left            =   8070
      TabIndex        =   10
      Top             =   1410
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
      MICON           =   "frm_list_PotonganLain.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   990
      TabIndex        =   1
      Top             =   930
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   195231747
      CurrentDate     =   40809
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1320
      OleObjectBlob   =   "frm_list_PotonganLain.frx":0070
      TabIndex        =   3
      Top             =   1440
      Width           =   1305
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   2670
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1440
      Width           =   3675
   End
   Begin VB.Frame Frame1 
      Height          =   8505
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   20025
      Begin prj_panji.LynxGrid LynxGrid1 
         Height          =   7275
         Left            =   3570
         TabIndex        =   9
         Top             =   210
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   12832
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
      Begin prj_panji.LynxGrid LynxGrid2 
         Height          =   7275
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   12832
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
      Left            =   3000
      TabIndex        =   2
      Top             =   930
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   161742851
      CurrentDate     =   40809
   End
   Begin prj_panji.vbButton btnImport 
      Height          =   345
      Left            =   6720
      TabIndex        =   16
      Top             =   1410
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Import Data"
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
      MICON           =   "frm_list_PotonganLain.frx":202E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      Caption         =   "* yyyy-MM-dd"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4680
      TabIndex        =   15
      Top             =   930
      Width           =   1665
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "LIST POTONGAN LAIN"
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
      Left            =   5850
      TabIndex        =   14
      Top             =   30
      Width           =   3435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "s/d"
      Height          =   195
      Left            =   2700
      TabIndex        =   7
      Top             =   990
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Periode :"
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   990
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Perusahaan :"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   1500
      Width           =   945
   End
End
Attribute VB_Name = "frm_List_PotonganLain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset
Dim strsql As String

Private Sub cmbshift_Change()
    LynxGrid1.Clear
End Sub

Private Sub btnImport_Click()
frm_import_other_expense.Show
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    LynxGrid1.Clear
    LynxGrid2.Clear
    isiGridDep
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    
    createGrid
    createGridDep
    
    Call isiCompany
    
    vbButton1.Enabled = False
    vbButton3.Enabled = False
    vbButton4.Enabled = False
    vbButton2.Enabled = False
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "Date", 1200, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Trans. Nr.", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "NIK", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Employee Name", 3500, , , , , , , , True
      .AddColumn "Div. Code", 1000, , , , , , , , False
      .AddColumn "Division", 2000, , , , , , , , , True
      .AddColumn "Title Code", 1000, , , , , , , , False
      .AddColumn "Title", 2000, , , , , , , , , True
      .AddColumn "Value", 1200, lgAlignCenterCenter, lgNumeric, "#,##", , , , , True
      .AddColumn "Remark", 1200, , , , , , , , True
      .AddColumn "Entry Date", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , True
      .AddColumn "Emp. Code", 1000, , , , , , , , False
      .AddColumn "Flag Other", 2500, , , , , , , , False
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

Public Sub isiGridAbsen()
Dim access As String

'access = IIf(LOGIN_LEVEL = 100, "", "AND (b.managerial_access = 0 OR b.managerial_access IS NULL)")

    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    strsql = "SELECT a.tgltrans,a.notrans," _
                & "b.nik,b.employee_name," _
                & "b.division_code,c.division_name," _
                & "b.title_code,d.title_name," _
                & "a.jmlPotong,a.remark, a.tglinput, a.employee_code, a.flag_other_expense " _
        & "FROM t_employee_expense a JOIN m_employee b ON a.employee_code = b.employee_code " _
        & "JOIN m_division c ON b.division_code = c.division_code and b.company_code = c.company_code " _
        & "JOIN m_title d ON b.title_code = d.title_code and b.company_code = d.company_code  " _
        & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " & access & " " _
        & "AND a.tgltrans BETWEEN '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker2.Value, "yyyy-MM-dd") & "' " _
        & "AND b.department_code = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "'"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!tgltrans & vbTab & rs!noTrans _
                & vbTab & rs!nik & vbTab & rs!EMPLOYEE_NAME _
                & vbTab & rs!division_code & vbTab & rs!division_name _
                & vbTab & rs!title_code & vbTab & rs!title_name _
                & vbTab & rs!jmlpotong & vbTab & rs!remark & vbTab & rs!tglinput _
                & vbTab & rs!employee_code & vbTab & rs!flag_other_expense
            rs.MoveNext
        Wend
    End If
    
    vbButton1.Enabled = IIf(rs.RecordCount = 0, False, True)
    vbButton3.Enabled = IIf(rs.RecordCount = 0, False, True)
    vbButton4.Enabled = IIf(rs.RecordCount = 0, False, True)
    vbButton2.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)
    
    rs.Close
    LynxGrid1.Redraw = True
End Sub

Private Sub showData()
    If LynxGrid1.Rows > 0 Then
        With frm_trans_PotongLain
            .TDBCombo_company.Text = TDBCombo_company.Text
            .txt_company_name.Text = txt_company_name.Text
            .cmbdep.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            .txt_dep_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            .DTPicker1.Value = LynxGrid1.CellText(LynxGrid1.Row, 0)
            .txtnotrans.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
            .txtkdkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 11)
            .txtnik.Text = LynxGrid1.CellText(LynxGrid1.Row, 2)
            .txtnmkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 3)
            .txtkddiv.Text = LynxGrid1.CellText(LynxGrid1.Row, 4)
            .txtdivision.Text = LynxGrid1.CellText(LynxGrid1.Row, 5)
            .txtkdtitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 6)
            .txttitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 7)
'            .txtjumlah.Value = LynxGrid1.CellValue(LynxGrid1.Row, 8)
'            .txtjumlah.Text = LynxGrid1.CellValue(LynxGrid1.Row, 8)
            .txtjumlah.Text = FormatNumber(DropAllComma(LynxGrid1.CellValue(LynxGrid1.Row, 8)))
            
            .Combo1.ListIndex = LynxGrid1.CellText(LynxGrid1.Row, 12)
            .txtket.Text = LynxGrid1.CellText(LynxGrid1.Row, 9)
            .v_value = .txtjumlah.Text
            
            .editTrans = True
            .Show vbModal
        End With
    End If
End Sub

Private Sub newData()
    With frm_trans_PotongLain
        .TDBCombo_company.Text = TDBCombo_company.Text
        .txt_company_name.Text = txt_company_name.Text
        .cmbdep.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
        .txt_dep_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
        .Show vbModal
    End With
End Sub

Private Sub Form_Resize()
    Frame1.Width = Me.Width - 500
    Frame1.Height = Me.Height - 2500
    LynxGrid1.Height = Frame1.Height - 400
    LynxGrid1.Width = Frame1.Width - (LynxGrid2.Width + 400)
    LynxGrid2.Height = Frame1.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_List_PotonganLain = Nothing
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

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
        
        Call isiGridDep
        LynxGrid1.Clear
    End If
End Sub

Private Sub vbButton1_Click()
    isiGridAbsen
End Sub

Private Sub vbbutton2_Click()
    newData
End Sub


Private Sub vbButton3_Click()
    showData
End Sub

Private Sub vbButton4_Click()
    Dim tanya As Integer
    tanya = MsgBox("Deleted Data Can Not be Undo, Are You Sure?", vbCritical + vbYesNo, "Warning")
    If tanya = vbYes Then
        CnG.Execute "DELETE FROM t_employee_expense where employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 11) & "' " _
            & "and notrans = '" & LynxGrid1.CellText(LynxGrid1.Row, 1) & "'"
        MsgBox "Delete Data Sucsess..."
        
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute strsql
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        isiGridAbsen
    End If
End Sub
