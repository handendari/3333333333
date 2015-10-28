VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_list_manual_overtime 
   Caption         =   "Manual Overtime"
   ClientHeight    =   10845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17280
   MDIChild        =   -1  'True
   ScaleHeight     =   10845
   ScaleWidth      =   17280
   WindowState     =   2  'Maximized
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   5460
      Top             =   690
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   2790
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   11
      Top             =   1260
      Width           =   3375
   End
   Begin prj_genting.vbButton vbButton3 
      Height          =   345
      Left            =   14220
      TabIndex        =   10
      Top             =   1230
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
      MICON           =   "frm_list_manual_overtime.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_genting.vbButton vbButton2 
      Height          =   345
      Left            =   12540
      TabIndex        =   9
      Top             =   1230
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
      MICON           =   "frm_list_manual_overtime.frx":001C
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
      Left            =   11100
      TabIndex        =   8
      Top             =   1230
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Add New Data"
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
      MICON           =   "frm_list_manual_overtime.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtnmshift 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   7950
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1260
      Width           =   2325
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   990
      TabIndex        =   1
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   118095875
      CurrentDate     =   40794
   End
   Begin VB.Frame Frame1 
      Height          =   7065
      Left            =   120
      TabIndex        =   0
      Top             =   1710
      Width           =   20025
      Begin prj_genting.LynxGrid LynxGrid2 
         Height          =   6615
         Left            =   120
         TabIndex        =   7
         Top             =   210
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   11668
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
      Begin prj_genting.LynxGrid LynxGrid1 
         Height          =   6615
         Left            =   3540
         TabIndex        =   6
         Top             =   210
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   11668
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
   Begin TrueOleDBList60.TDBCombo TDBCombo1 
      Height          =   375
      Left            =   7020
      OleObjectBlob   =   "frm_list_manual_overtime.frx":0054
      TabIndex        =   2
      Top             =   1260
      Width           =   885
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   990
      OleObjectBlob   =   "frm_list_manual_overtime.frx":1FFF
      TabIndex        =   12
      Top             =   1260
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1350
      Top             =   1470
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin prj_genting.vbButton vbButton4 
      Height          =   345
      Left            =   2820
      TabIndex        =   14
      Top             =   840
      Width           =   1395
      _ExtentX        =   2461
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
      MICON           =   "frm_list_manual_overtime.frx":3FBD
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
      Left            =   1020
      TabIndex        =   16
      Top             =   570
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "LIST OVERTIME"
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
      Left            =   6150
      TabIndex        =   15
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Company"
      Height          =   195
      Left            =   300
      TabIndex        =   13
      Top             =   1350
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Shift"
      Height          =   195
      Left            =   6480
      TabIndex        =   4
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "Date :"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   900
      Width           =   495
   End
End
Attribute VB_Name = "frm_list_manual_overtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim strsql As String
Dim v_15, v_2, v_3, v_4 As Double
Dim v_tot_ot, v_holiday As Double

Private Sub cmbshift_Change()
    LynxGrid1.Clear
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    LynxGrid1.Clear
    LynxGrid2.Clear
    isiGridDep
End Sub

Private Sub Form_Load()
    Adodc_company.ConnectionString = strConn
    DTPicker1.Value = Date
    
    Call load_data_company
    
    createGrid
    createGridDep
    
    timer1.Enabled = True
    
    vbButton2.Enabled = False
    vbButton3.Enabled = False
    vbButton4.Enabled = False
    vbButton1.Enabled = False
    
'    Call isiCompany
'    Call isiAbsentStatus
End Sub

Private Sub isiShift()
Dim rsshift As New ADODB.Recordset
    strsql = "select shift_code,shift_name " _
            & "from m_shift where company_code = '" & TDBCombo_company.Text & "'"
    rsshift.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo1.RowSource = rsshift
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "NIK", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Employee Name", 3500, , , , , , , , True
      .AddColumn "Div. Code", 1000, , , , , , , , False
      .AddColumn "Division", 1700, , , , , , , , , True
      .AddColumn "Title Code", 1000, , , , , , , , False
      .AddColumn "Title", 1500, , , , , , , , , True
      .AddColumn "Time IN", 800, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
      .AddColumn "Time OUT", 900, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
      .AddColumn "Tot. Hour", 1000, lgAlignRightCenter, lgNumeric, "#.##,##", , , , , False
      .AddColumn "1.5", 500, lgAlignRightCenter, lgNumeric, "#.##,##", , , , , False
      .AddColumn "2", 500, lgAlignRightCenter, lgNumeric, "#.##,##", , , , , False
      .AddColumn "3", 500, lgAlignRightCenter, lgNumeric, "#.##,##", , , , , False
      .AddColumn "4", 500, lgAlignRightCenter, lgNumeric, "#.##,##", , , , , False
      .AddColumn "holiday", , , lgNumeric, , , , , , False
      .AddColumn "Entry Date", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , True
      .AddColumn "Employee Code", , , , , , , , , False
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
    Dim kdshift As String
    
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    strsql = "select b.nik,b.employee_name," _
                & "b.division_code,c.division_name,b.title_code,d.title_name," _
                & "a.total_hour,a.15jam,a.2jam,a.3jam,a.4jam,holiday,a.entry_date,b.employee_code,a.time_in,a.time_out " _
        & "from m_employee b LEFT JOIN h_overtime a on b.employee_code = a.employee_code " _
        & "left join m_department e on e.department_code = b.department_code and e.company_code = b.company_code " _
        & "left join m_division c on b.division_code = c.division_code and b.department_code = c.department_code and b.company_code = c.company_code " _
        & "left join m_title d on b.title_code = d.title_code and d.company_code = b.company_code " _
        & "WHERE a.att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " _
        & "AND a.shift_code = '" & TDBCombo1.Text & "' " _
        & "AND b.department_code = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' " _
        & "AND (b.level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
'            LynxGrid1.AddItem rs!nik & vbTab & rs!employee_name _
'                & vbTab & IIf(IsNull(rs!division_code), "", rs!division_code) & vbTab & IIf(IsNull(rs!division_name), "", rs!division_name) _
'                & vbTab & rs!title_code & vbTab & rs!title_name _
'                & vbTab & rs!total_hour & vbTab & rs.Fields("15jam") & vbTab & rs.Fields("2jam") & vbTab & rs.Fields("3jam") _
'                & vbTab & rs.Fields("4jam") & vbTab & rs!holiday & vbTab & rs!entry_date & vbTab & rs!employee_code

            LynxGrid1.AddItem rs!nik & vbTab & rs!employee_name _
                & vbTab & IIf(IsNull(rs!division_code), "", rs!division_code) & vbTab & IIf(IsNull(rs!division_name), "", rs!division_name) _
                & vbTab & rs!title_code & vbTab & rs!title_name _
                & vbTab & rs!time_in & vbTab & rs!time_out _
                & vbTab & rs!total_hour & vbTab & rs.Fields("15jam") & vbTab & rs.Fields("2jam") & vbTab & rs.Fields("3jam") _
                & vbTab & rs.Fields("4jam") _
                & vbTab & rs!holiday & vbTab & rs!entry_date & vbTab & rs!employee_code
            rs.MoveNext
        Wend
    End If
    
    vbButton2.Enabled = IIf(rs.RecordCount = 0, False, True)
    vbButton3.Enabled = IIf(rs.RecordCount = 0, False, True)
    vbButton4.Enabled = IIf(rs.RecordCount = 0, False, True)
    vbButton1.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)
    
    rs.Close
    LynxGrid1.Redraw = True
End Sub

Private Sub showData()
Dim v_holiday As Integer

strsql = "SELECT holiday FROM h_overtime WHERE " _
        & "employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 15) & "' AND " _
        & "att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
    v_holiday = rs!holiday
End If
rs.Close

    If LynxGrid1.Rows > 0 Then
        With frm_trans_overtime
            .DTPicker1.Value = DTPicker1.Value
            .txtkdshift.Text = TDBCombo1.Text
            .txtnmshift.Text = txtnmshift.Text
            .cmbdep.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            .txtnmDept.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            .txtkdkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 15)
            .txt_nik.Text = LynxGrid1.CellText(LynxGrid1.Row, 0)
            .txtnmkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
            .txtkddiv.Text = LynxGrid1.CellText(LynxGrid1.Row, 2)
            .txtdivision.Text = LynxGrid1.CellText(LynxGrid1.Row, 3)
            .txtkdtitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 4)
            .txttitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 5)
            .txtjmljam.Text = LynxGrid1.CellValue(LynxGrid1.Row, 8)
            .txt15jam.Text = LynxGrid1.CellValue(LynxGrid1.Row, 9)
            .txt2jam.Text = LynxGrid1.CellValue(LynxGrid1.Row, 10)
            .Txt3jam.Text = LynxGrid1.CellValue(LynxGrid1.Row, 11)
            .Txt4jam.Text = LynxGrid1.CellText(LynxGrid1.Row, 12)
            .Check1.Value = LynxGrid1.CellValue(LynxGrid1.Row, 13)
            .ttin.Value = Format(LynxGrid1.CellText(LynxGrid1.Row, 6), "hh:mm")
            .ttout.Value = Format(LynxGrid1.CellText(LynxGrid1.Row, 7), "hh:mm")
'            .TDBCombo1.BoundText = TDBCombo2.Text
'            .txtnmabsentstatus.Text = txtabsentname.Text
            
            .v_ot_hour = .txtjmljam.Text
            .v_tt_in = Format(LynxGrid1.CellText(LynxGrid1.Row, 6), "hh:mm")
            .v_tt_out = Format(LynxGrid1.CellText(LynxGrid1.Row, 7), "hh:mm")
            
            .Check1.Value = v_holiday
            
            .editTrans = True
            .Show vbModal
        End With
    End If
End Sub

Private Sub newData()
    With frm_trans_overtime
        .DTPicker1.Value = DTPicker1.Value
        .txtkdshift.Text = TDBCombo1.Text
        .txtnmshift.Text = txtnmshift.Text
        .cmbdep.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
        .txtnmDept.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
        .txt15jam.Text = "0"
        .txt2jam.Text = "0"
        .Txt3jam.Text = "0"
        .Txt4jam.Text = "0"
        .editTrans = False
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
    Set frm_list_manual_att = Nothing
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
        Call isiShift
        LynxGrid1.Clear
    End If
End Sub

Private Sub TDBCombo1_ItemChange()
    If TDBCombo1.ApproxCount > 0 Then
        TDBCombo1.Text = TDBCombo1.Columns("shift_code").Value
        txtnmshift.Text = TDBCombo1.Columns("shift_name").Value
        If LynxGrid2.Rows > 0 Then
            isiGridAbsen
        End If
    End If
End Sub

Private Sub TDBCombo2_ItemChange()
    If TDBCombo2.ApproxCount > 0 Then
        TDBCombo2.Text = TDBCombo2.Columns("absent_code").Value
        txtabsentname.Text = TDBCombo2.Columns("absent_name").Value
        LynxGrid1.Clear
        If LynxGrid2.Rows > 0 Then
            isiGridAbsen
        End If
        If TDBCombo2.Text <> "0" Then
            vbButton2.Enabled = False
        Else
            vbButton2.Enabled = True
        End If
    End If
End Sub

Private Sub vbButton1_Click()
    If LynxGrid2.Rows = 0 Then
        MsgBox "Invalid Departement Code." & Chr(13) & "Please Check Your Transaction Again...", vbInformation, "Error"
        Exit Sub
    End If
    If TDBCombo1.Text = "" Then
        MsgBox "Invalid Shift Code." & Chr(13) & "Please Check Your Transaction Again...", vbInformation, "Error"
        Exit Sub
    End If
    newData
End Sub

Private Sub vbButton2_Click()
    showData
End Sub

Private Sub vbButton3_Click()
Dim tanya As Integer
Dim tot_hours As Double
Dim tt_in As String
Dim tt_out As String
Dim tot_ot As Double

Dim rsot As New ADODB.Recordset
        
tt_in = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(LynxGrid1.CellText(LynxGrid1.Row, 6), "hh:mm") & ":00"
tt_out = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(LynxGrid1.CellText(LynxGrid1.Row, 7), "hh:mm") & ":00"

        strsql = "SELECT total_hour FROM h_overtime " _
            & "WHERE employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 15) & "' " _
            & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " _
            & "AND time_in = '" & tt_in & "' " _
            & "AND time_out = '" & tt_out & "'"
        rsot.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsot.RecordCount > 0 Then
            tot_hours = rsot!total_hour
        End If
        rsot.Close
        
        strsql = "SELECT total_ot,holiday FROM h_attendance " _
        & "WHERE employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 15) & "' " _
        & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
        rsot.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsot.RecordCount > 0 Then
            v_tot_ot = rsot!total_OT
            v_holiday = rsot!holiday
        End If
        rsot.Close

        v_tot_ot = (v_tot_ot - tot_hours)

        Call totalOT
        tot_ot = ((1.5 * Val(v_15)) + (2 * Val(v_2)) + (3 * Val(v_3)) + (4 * Val(v_4)))
        
        tanya = MsgBox("Deleted Data Can Not be Undo..!!", vbCritical + vbYesNo, "Warning")
        If tanya = vbYes Then
            CnG.Execute "DELETE FROM h_overtime where employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 15) & "' " _
                & "and att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' and shift_code = '" & TDBCombo1.Text & "' " _
                & "AND time_in = '" & tt_in & "' " _
                & "AND time_out = '" & tt_out & "'"
            
            '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
            strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
            CnG.Execute strsql
            
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            strsql = "UPDATE h_attendance set " _
                & "att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',shift_code = '" & TDBCombo1.Text & "'," _
                & "total_ot = '" & v_tot_ot & "',15jam = '" & v_15 & "'," _
                & "2jam = '" & v_2 & "',3jam = '" & v_3 & "'," _
                & "4jam = '" & v_4 & "',holiday = '" & v_holiday & "'," _
                & "tot_overtime = '" & tot_ot & "', editdate=now() " _
                & "WHERE employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 15) & "' " _
                & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
            CnG.Execute strsql
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            MsgBox "Delete Data Sucsess..."
            isiGridAbsen
        End If
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub vbButton4_Click()
    isiGridAbsen
End Sub

Private Sub Timer1_Timer()
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Public Sub totalOT()
If v_holiday = 1 Then
    v_tot_work = 0
    If v_tot_ot > 15 Then
        v_tot_ot = 15
    Else
        v_tot_ot = v_tot_ot
    End If
        Select Case Val(v_tot_ot)
            Case Is <= 7
                v_15 = 0
                v_2 = Val(v_tot_ot)
                v_3 = 0
                v_4 = 0
            Case Is = 7.5
                v_15 = 0
                v_2 = 7
                v_3 = 0.5
                v_4 = 0
            Case Is = 8
                v_15 = 0
                v_2 = 7
                v_3 = 1
                v_4 = 0
            Case Else
                v_15 = 0
                v_2 = 7
                v_3 = 1
                v_4 = Val(v_tot_ot) - Val(v_2) - Val(v_3)
        End Select
    Else
        Select Case Val(v_tot_ot)
            Case Is <= 1
                v_15 = 1
                v_2 = 0
                v_3 = 0
                v_4 = 0
            Case Is > 1
                v_15 = 1
                v_2 = Val(v_tot_ot) - Val(v_15)
                v_3 = 0
                v_4 = 0
            Case Else
                v_15 = 0
                v_2 = 0
                v_3 = 0
                v_4 = 0
        End Select
    End If
End Sub
