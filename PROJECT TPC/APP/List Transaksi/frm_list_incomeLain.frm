VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_List_IncomeLain 
   Caption         =   "OTHER INCOME / EXPENSE"
   ClientHeight    =   10140
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   Icon            =   "frm_list_incomeLain.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10140
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fra_status_emp 
      Caption         =   "List Income Expense"
      Height          =   585
      Left            =   9240
      TabIndex        =   14
      Top             =   720
      Width           =   2925
      Begin VB.OptionButton optIncome 
         Caption         =   "Income"
         Height          =   225
         Left            =   240
         TabIndex        =   16
         Top             =   270
         Width           =   855
      End
      Begin VB.OptionButton optExpense 
         Caption         =   "Expense"
         Height          =   225
         Left            =   1500
         TabIndex        =   15
         Top             =   270
         Width           =   1245
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Top             =   1080
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   94568451
      CurrentDate     =   40809
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1170
      OleObjectBlob   =   "frm_list_incomeLain.frx":058A
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   2430
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   1950
      Width           =   20025
      Begin prj_tpc.LynxGrid LynxGrid1 
         Height          =   7515
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   11895
         _ExtentX        =   20981
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
      Left            =   2790
      TabIndex        =   2
      Top             =   1080
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   94568451
      CurrentDate     =   40809
   End
   Begin prj_tpc.vbButton cmdDelete 
      Height          =   465
      Left            =   10830
      TabIndex        =   10
      Top             =   1350
      Width           =   1305
      _ExtentX        =   2302
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
      MICON           =   "frm_list_incomeLain.frx":24F0
      PICN            =   "frm_list_incomeLain.frx":250C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdEdit 
      Height          =   465
      Left            =   9480
      TabIndex        =   11
      Top             =   1350
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&Edit"
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
      MICON           =   "frm_list_incomeLain.frx":359E
      PICN            =   "frm_list_incomeLain.frx":35BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdNew 
      Height          =   465
      Left            =   8130
      TabIndex        =   12
      Top             =   1350
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&New"
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
      MICON           =   "frm_list_incomeLain.frx":464C
      PICN            =   "frm_list_incomeLain.frx":4668
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdRefresh 
      Height          =   495
      Left            =   4650
      TabIndex        =   13
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
      MICON           =   "frm_list_incomeLain.frx":56FA
      PICN            =   "frm_list_incomeLain.frx":5716
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker_Periode 
      Height          =   315
      Left            =   1170
      TabIndex        =   17
      Top             =   720
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MM-yyyy"
      Format          =   94568451
      CurrentDate     =   40794
   End
   Begin VB.Label Label5 
      Caption         =   "PERIODE :"
      Height          =   195
      Left            =   300
      TabIndex        =   18
      Top             =   780
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "OTHER INCOME / EXPENSE"
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
      TabIndex        =   9
      Top             =   150
      Width           =   3945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TO"
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
      Left            =   2490
      TabIndex        =   7
      Top             =   1140
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATE :"
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   1140
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY :"
      Height          =   195
      Left            =   210
      TabIndex        =   5
      Top             =   1500
      Width           =   885
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_list_incomeLain.frx":67A8
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
Dim rsCompany As New ADODB.Recordset
    
Private Sub DTPicker_Periode_Change()
    Call getPeriode(DTPicker_Periode.Value, DTPicker1, DTPicker2)
    
    Call isiGridAbsen
End Sub

'Private Sub btnImport_Click()
'frm_import_other_income.Show
'End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    LynxGrid1.Clear
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker_Periode.Value = Date
    
    createGrid
    
    Timer1.Enabled = True
    Call isiCompany
    
    cmdNew.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "DATE", 1200, lgAlignCenterCenter, , , , , , , True
      .AddColumn "EMP. CODE", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "EMP. NAME", 2500, , , , , , , , True
      .AddColumn "DEPT. CODE", , , , , , , , , False
      .AddColumn "DEPARTMENT", 1500, , , , , , , , , True
      .AddColumn "DIV. CODE", , , , , , , , , False
      .AddColumn "SECTION", 1500, , , , , , , , , True
      .AddColumn "TITLE CODE", , , , , , , , , False
      .AddColumn "JOB TITLE", 2000, , , , , , , , , True
      .AddColumn "TYPE", 2000, lgAlignCenterCenter, , , , , , , , True
      .AddColumn "VALUE", 1200, lgAlignCenterCenter, lgNumeric, "#,##", , , , , True
      .AddColumn "DESCRIPTION", 1500, , , , , , , , True
      .AddColumn "ENTRY DATE", 1500, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , True
      .AddColumn "Employee Code", 2500, , , , , , , , False
      .AddColumn "Flag Income Expense", 1500, , , , , , , , False
      .AddColumn "Nomer", 2500, , , , , , , , False
      .AddColumn "Flag Type", 2500, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiCompany()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo_company.RowSource = rsCompany
End Sub

Public Sub isiGridAbsen()
Dim access As String

    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    If optIncome.Value = True Then
        SQL = "SELECT a.tgltrans,b.nik,b.employee_name,b.department_code,e.department_name," _
                    & "b.division_code,c.division_name," _
                    & "b.title_code,d.title_name,CASE WHEN flag_income_expense = 1 THEN '' ELSE nm_type END nm_type," _
                    & "a.jmlPotong,a.remark, a.tglinput, b.employee_code, " _
                    & "a.flag_income_expense, a.nomer,a.flag_type " _
            & "FROM t_income_expense a JOIN m_employee b ON a.employee_code = b.employee_code " _
            & "JOIN m_department e ON b.department_code = e.department_code " _
            & "JOIN m_division c ON b.division_code = c.division_code and b.company_code = c.company_code " _
            & "JOIN m_title d ON b.title_code = d.title_code " _
            & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " _
            & "AND a.tgltrans BETWEEN '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker2.Value, "yyyy-MM-dd") & "' " _
            & "AND a.flag_income_expense = 0"
    Else
        SQL = "SELECT a.tgltrans,b.nik,b.employee_name,b.department_code,e.department_name," _
                    & "b.division_code,c.division_name," _
                    & "b.title_code,d.title_name,CASE WHEN flag_income_expense = 1 THEN '' ELSE nm_type END nm_type," _
                    & "a.jmlPotong,a.remark, a.tglinput, b.employee_code, " _
                    & "a.flag_income_expense, a.nomer,a.flag_type " _
            & "FROM t_income_expense a JOIN m_employee b ON a.employee_code = b.employee_code " _
            & "JOIN m_department e ON b.department_code = e.department_code " _
            & "JOIN m_division c ON b.division_code = c.division_code and b.company_code = c.company_code " _
            & "JOIN m_title d ON b.title_code = d.title_code " _
            & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " _
            & "AND a.tgltrans BETWEEN '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker2.Value, "yyyy-MM-dd") & "' " _
            & "AND a.flag_income_expense = 1"
    End If
    
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!tgltrans & vbTab & rs!nik & vbTab & rs!EMPLOYEE_NAME _
                & vbTab & rs!DEPARTMENT_CODE & vbTab & rs!department_name _
                & vbTab & rs!DIVISION_CODE & vbTab & rs!division_name _
                & vbTab & rs!title_code & vbTab & rs!title_name & vbTab & rs!nm_type _
                & vbTab & rs!jmlpotong & vbTab & rs!remark _
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
    
    Call load_data_user_access(Me)
    cmdNew.Enabled = blnUser_Add
    cmdEdit.Enabled = blnUser_Edit
    cmdDelete.Enabled = blnUser_Delete
End Sub

Private Sub showData()
    If LynxGrid1.Rows > 0 Then
        With frm_trans_IncomeLain
            .TDBCombo_company.Text = TDBCombo_company.Text
            .txt_company_name.Text = txt_company_name.Text
            .DTPicker1.Value = LynxGrid1.CellText(LynxGrid1.Row, 0)
            .txtkdkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 13)
            .txtnik.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
            .txtnmkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 2)
            .txt_department_code.Text = LynxGrid1.CellText(LynxGrid1.Row, 3)
            .txt_dep_name.Text = LynxGrid1.CellText(LynxGrid1.Row, 4)
            .txtkddiv.Text = LynxGrid1.CellText(LynxGrid1.Row, 5)
            .txtdivision.Text = LynxGrid1.CellText(LynxGrid1.Row, 6)
            .txtkdtitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 7)
            .txttitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 8)
            .txtjumlah.Text = FormatNumber(DropAllComma(LynxGrid1.CellValue(LynxGrid1.Row, 10)))
            
            .v_value = .txtjumlah.Text
            .v_nomer = LynxGrid1.CellText(LynxGrid1.Row, 15)
            .Combo1.ListIndex = LynxGrid1.CellText(LynxGrid1.Row, 14)
            .cboType.ListIndex = IIf(IsNull(LynxGrid1.CellText(LynxGrid1.Row, 16)), 0, LynxGrid1.CellText(LynxGrid1.Row, 16))
            .cboType.Enabled = False
            
            If .Combo1.ListIndex = 0 Then
                .cboType.Visible = True
            Else
                .cboType.Visible = False
            End If
            
            
            .txtket.Text = LynxGrid1.CellText(LynxGrid1.Row, 11)
            .editTrans = True
            .Show vbModal
        End With
    End If
End Sub

Private Sub newData()
    With frm_trans_IncomeLain
        .TDBCombo_company.Text = TDBCombo_company.Text
        .txt_company_name.Text = txt_company_name.Text
        .Combo1.ListIndex = IIf(optIncome.Value, 0, 1)
        .cboType.ListIndex = 0
        
        If .Combo1.ListIndex = 0 Then
            .cboType.Visible = True
        Else
            .cboType.Visible = False
        End If
'        If optIncome.Value Then
'            .Combo2.Enabled = True
'            .Combo2.Visible = True
'            .Combo2.ListIndex = 0
'        Else
'            .Combo2.Visible = False
'        End If
        .Show vbModal
    End With
End Sub

Private Sub Form_Resize()
    Frame1.Width = Me.Width - 500
    Frame1.Height = Me.Height - 2500
    LynxGrid1.Height = Frame1.Height - 400
    LynxGrid1.Width = Frame1.Width + 400
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
        
        Call isiGridAbsen
        optIncome.Value = True
    End If
End Sub

Private Sub cmdRefresh_Click()
    isiGridAbsen
End Sub

Private Sub CmdNew_Click()
    newData
End Sub


Private Sub cmdEdit_Click()
    showData
End Sub

Private Sub cmdDelete_Click()
    Dim tanya As Integer
    tanya = MsgBox("Deleted Data Can Not be Undo, Are You Sure?", vbCritical + vbYesNo, "Warning")
    If tanya = vbYes Then
        CnG.Execute "DELETE FROM t_income_expense where employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 13) & "' " _
            & "and nomer = '" & LynxGrid1.CellText(LynxGrid1.Row, 15) & "' " _
            & "AND flag_income_expense = '" & IIf(optIncome.Value, 0, 1) & "'"
        MsgBox "Delete Data Sucsess..."
        
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute SQL
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        isiGridAbsen
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub
