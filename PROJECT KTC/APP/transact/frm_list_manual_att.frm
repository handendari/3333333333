VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_list_manual_att 
   Caption         =   "Manual Attendance"
   ClientHeight    =   10755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17280
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10755
   ScaleWidth      =   17280
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   5580
      Top             =   150
   End
   Begin prj_fej_jkt.vbButton cmdBack 
      Height          =   315
      Left            =   2700
      TabIndex        =   15
      Top             =   210
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "<"
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
      MICON           =   "frm_list_manual_att.frx":0000
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
      Left            =   9300
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   12
      Top             =   600
      Width           =   2325
   End
   Begin prj_fej_jkt.vbButton vbButton1 
      Height          =   345
      Left            =   11970
      TabIndex        =   4
      Top             =   570
      Width           =   1275
      _ExtentX        =   2249
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
      MICON           =   "frm_list_manual_att.frx":001C
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
      OleObjectBlob   =   "frm_list_manual_att.frx":0038
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
      TabIndex        =   7
      Top             =   600
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   810
      TabIndex        =   1
      Top             =   210
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   94896131
      CurrentDate     =   40794
   End
   Begin prj_fej_jkt.vbButton vbButton2 
      Height          =   345
      Left            =   13290
      TabIndex        =   5
      Top             =   570
      Width           =   1350
      _ExtentX        =   2381
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
      MICON           =   "frm_list_manual_att.frx":1F9E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo1 
      Height          =   375
      Left            =   8370
      OleObjectBlob   =   "frm_list_manual_att.frx":1FBA
      TabIndex        =   3
      Top             =   600
      Width           =   885
   End
   Begin prj_fej_jkt.vbButton vbButton3 
      Height          =   345
      Left            =   14940
      TabIndex        =   13
      Top             =   570
      Width           =   1245
      _ExtentX        =   2196
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
      MICON           =   "frm_list_manual_att.frx":3F0D
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
      Height          =   315
      Left            =   3780
      TabIndex        =   14
      Top             =   210
      Width           =   1275
      _ExtentX        =   2249
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
      MICON           =   "frm_list_manual_att.frx":3F29
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
      Height          =   9315
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   20025
      Begin prj_fej_jkt.LynxGrid LynxGrid2 
         Height          =   8895
         Left            =   120
         TabIndex        =   10
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
         Left            =   3570
         TabIndex        =   6
         Top             =   210
         Width           =   12390
         _ExtentX        =   21855
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
         ColumnSort      =   -1  'True
      End
   End
   Begin prj_fej_jkt.vbButton cmdNext 
      Height          =   315
      Left            =   3150
      TabIndex        =   16
      Top             =   210
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   ">"
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
      MICON           =   "frm_list_manual_att.frx":3F45
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Work Status :"
      Height          =   195
      Left            =   7350
      TabIndex        =   11
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Date :"
      Height          =   195
      Left            =   300
      TabIndex        =   9
      Top             =   270
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Branch Office :"
      Height          =   195
      Left            =   300
      TabIndex        =   8
      Top             =   660
      Width           =   1065
   End
End
Attribute VB_Name = "frm_list_manual_att"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim strsql As String
Dim rscompany As New ADODB.Recordset

Private Sub cmbshift_Change()
    LynxGrid1.Clear
End Sub

Private Sub cmdBack_Click()
DTPicker1.Value = DTPicker1 - 1

Call isiGridAbsen
End Sub

Private Sub cmdNext_Click()
DTPicker1.Value = DTPicker1 + 1

Call isiGridAbsen
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    LynxGrid1.Clear
    LynxGrid2.Clear
    isiGridDep
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Format(Now, "yyyy-MM-dd")
    
    createGrid
    createGridDep
    
    Call isiShift
    Call isiCompany
    'Call isiAbsentStatus
    
    vbButton2.Enabled = False
    vbButton3.Enabled = False
    vbButton4.Enabled = False
    vbButton1.Enabled = False
    
    timer1.Enabled = True
End Sub

Private Sub isiShift()
Dim rsshift As New ADODB.Recordset
    strsql = "select shift_code,shift_name " _
            & "from m_shift"
    rsshift.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo1.RowSource = rsshift
End Sub

Private Sub isiAbsentStatus()
Dim rsabs As New ADODB.Recordset
    strsql = "select absent_code,absent_name " _
            & "from m_absent_status"
    rsabs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo2.RowSource = rsabs
    TDBCombo2.BoundText = "0"
    txtabsentname.Text = "PRESENT"
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "NIK", 1300, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Employee Name", 2500, , , , , , , , True
'      .BackColorSel = &H8000000D
      .AddColumn "Shift", 800, lgAlignCenterCenter, , , , , , , , True
      .AddColumn "Title Code", 1000, , , , , , , , False
      .AddColumn "Title", 2000, , , , , , , , , True
      .AddColumn "Cek IN", 1000, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
      .AddColumn "Break OUT", 1000, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
      .AddColumn "Break IN", 1000, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
      .AddColumn "Cek OUT", 1000, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
      .AddColumn "Total Hours", 1000, lgAlignCenterCenter, lgNumeric, "#.##", , , , , True
      .AddColumn "Entry Date", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub createGridDep()
   With LynxGrid2
      .AddColumn "Code", 800, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Group", 2500, , , , , , , , True
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

'Private Sub isiGridDep()
'    LynxGrid2.Redraw = False
'    LynxGrid2.Clear
'    strsql = "select department_code,department_name " _
'            & "from m_department where company_code = '" & TDBCombo_company.Text & "'"
'    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'    If rs.RecordCount > 0 Then
'        rs.MoveFirst
'        While Not rs.EOF
'            LynxGrid2.AddItem rs!department_code & vbTab & rs!department_name
'            rs.MoveNext
'        Wend
'    End If
'    rs.Close
'    LynxGrid2.Redraw = True
'End Sub

Private Sub isiGridDep()
    LynxGrid2.Redraw = False
    LynxGrid2.Clear
    strsql = "select group_code,group_name " _
            & "from m_group_shift where company_code = '" & TDBCombo_company.Text & "'"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid2.AddItem rs!group_code & vbTab & rs!group_name
            rs.MoveNext
        Wend
    End If
    rs.Close
    LynxGrid2.Redraw = True
End Sub

Public Sub isiGridAbsen()
    Dim kdshift As String
    Dim access As String
    
    access = IIf(LOGIN_LEVEL = 100, "", "AND (b.managerial_access = 0 OR b.managerial_access IS NULL)")
    
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    If TDBCombo1.Text = "ALL" Then
        strsql = "select a.employee_code,b.employee_name," _
                    & "a.shift_code,b.title_code,d.title_name," _
                    & "a.time_in,a.time_out_break,a.time_in_break,a.time_out,b.department_code," _
                    & "fn_get_jml_jam_kerja(a.time_out,a.time_in,a.time_in_break,a.time_out_break) tot_hours,a.entry_date " _
            & "from m_employee b LEFT JOIN h_attendance a on b.employee_code = a.employee_code " _
            & "join m_title d on b.title_code = d.title_code " _
            & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " _
            & "AND date(a.att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " _
            & "AND b.group_shift = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' " _
            & "AND b.flag_active = 1 " & access & " " _
            & "ORDER BY b.employee_name"
    Else
        strsql = "select a.employee_code,b.employee_name," _
                    & "a.shift_code,b.title_code,d.title_name," _
                    & "a.time_in,a.time_out_break,a.time_in_break,a.time_out,b.department_code," _
                    & "fn_get_jml_jam_kerja(a.time_out,a.time_in,a.time_in_break,a.time_out_break) tot_hours,a.entry_date " _
            & "from m_employee b LEFT JOIN h_attendance a on b.employee_code = a.employee_code " _
            & "join m_title d on b.title_code = d.title_code " _
            & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " _
            & "AND a.shift_code = '" & TDBCombo1.Text & "' " _
            & "AND date(a.att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " _
            & "AND b.group_shift = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "' " _
            & "AND b.flag_active = 1 " & access & " " _
            & "ORDER BY b.employee_name"
    End If
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!employee_code & vbTab & rs!employee_name _
                & vbTab & rs!shift_code & vbTab & rs!title_code & vbTab & rs!title_name _
                & vbTab & rs!time_in & vbTab & rs!time_out_break & vbTab & rs!time_in_break & vbTab & rs!time_out _
                & vbTab & rs!tot_hours & vbTab & rs!entry_date
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
Dim break As String

'On Error GoTo err

    If LynxGrid1.Rows > 0 Then
        Dim rs As New ADODB.Recordset
        strsql = "select a.att_date,c.employee_code,c.employee_name," & _
                    "c.division_code,d.division_name,c.department_code,e.department_name," & _
                    "c.title_code,f.title_name,a.att_date,a.shift_code," & _
                    "case when b.shift_name is null then 'N' else b.shift_code end shift_code," & _
                    "case when b.shift_name is null then 'NORMAL' else b.shift_name end shift_name," & _
                    "(SELECT shift_name FROM m_shift WHERE shift_code = a.shift_number) nama_shift," & _
                    "case when b.shift_name is null then a.shift_code else 0 end absent_code," & _
                    "CASE WHEN b.shift_name is null then " & _
                       "(select absent_name from m_absent_status where absent_code = a.shift_code) " & _
                    "else b.shift_name end absent_name," & _
                    "a.start_time,a.end_time,a.time_in,a.time_out," & _
                    "CAST(LEFT(TIMEDIFF(a.time_in_break,a.time_out_break),2) AS Char) break, " & _
                    "a.absent_status,a.entry_date,c.company_code,g.company_name,a.description,a.shift_number " & _
            "from h_attendance a join m_employee c on a.employee_code = c.employee_code " & _
            "join m_division d on c.division_code = d.division_code and c.company_code = d.company_code " & _
            "join m_department e on c.department_code = e.department_code and c.company_code = e.company_code " & _
            "join m_title f on c.title_code = f.title_code " & _
            "join m_company g on c.company_code = g.company_code " & _
            "left join m_shift b on a.shift_code = b.shift_code " & _
            "left join m_absent_status h on a.shift_code = h.absent_code " & _
            "WHERE date(a.att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' AND a.employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 0) & "'"
        
'"case when b.shift_name is null then 'NORMAL' else b.shift_name end shift_name," &

        rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
            
            break = IIf(IsNull(rs!break), "00", rs!break)
        
        With frm_trans_att_man
            .isijamkerja
            .DTPicker1.Value = rs!att_date
            .TDBCombo_company.Text = rs!COMPANY_CODE
            .txt_company_name.Text = rs!company_name
'            .txtkdshift.Text = rs!shift_code
            If rs!shift_number = 0 Or rs!shift_number = 1 Or IsNull(rs!shift_number) Then
                .TDBCombo2.Text = rs!shift_code
                .TDBCombo1.Text = rs!shift_code
                .txtnmshift.Text = rs!shift_name
            Else
                .TDBCombo2.Text = rs!shift_number
                .TDBCombo1.Text = rs!shift_code
                .txtnmshift.Text = rs!nama_shift
            End If
            .cmbdep.Text = rs!department_code
            .txtnmDept.Text = rs!department_name
            .txtkdkar.Text = rs!employee_code
            .txtnmkar.Text = rs!employee_name
            .txtkddiv.Text = rs!division_code
            .txtdivision.Text = rs!division_name
            .txtkdtitle.Text = rs!title_code
            .txttitle.Text = rs!title_name
            .ttin.Value = Format(rs!time_in, "hh:mm")
            .ttout.Value = Format(rs!time_out, "hh:mm")
            .txtentry.Text = rs!entry_date
            .TDBCombo1.BoundText = IIf(IsNull(rs!absent_code), 0, rs!absent_code)
            If break = "02" Then
                .Combo1.Text = "2 Jam Break"
            ElseIf break = "01" Then
                .Combo1.Text = "1 Jam Break"
            Else
                .Combo1.Text = "0 Jam Break"
            End If
            
            If .TDBCombo1.Text <> "0" Then
                .Combo3.Visible = False
                .DTPicker2.Value = rs!att_date
                .DTPicker2.Visible = True
                .ttin.Enabled = False
                .ttout.Enabled = False
                .Combo1.ListIndex = 0
                .Combo1.Enabled = False
                '.Label16.Visible = True
            Else
                .Combo3.Visible = False
                .DTPicker2.Value = rs!att_date
                .DTPicker2.Visible = True
                .ttin.Enabled = True
                .ttout.Enabled = True
'                .Combo1.Enabled = False
            End If
            
            If .TDBCombo1.Text = "0" Then
                .txtnmabsentstatus.Text = "PRESENT"
            Else
                .txtnmabsentstatus.Text = IIf(IsNull(rs!absent_name), "MEDICAL CHECK", rs!absent_name)
            End If
                
            .txtket.Text = IIf(IsNull(rs!Description), "", rs!Description)
            .chk_all_employee.Enabled = False
            
            .v_absen_status = .TDBCombo1.Text
            .v_tt_in = .ttin.Value
            .v_tt_out = .ttout.Value
            
            .editTrans = True
            .Show vbModal
        End With
        rs.Close
    End If
    Exit Sub

err:
MsgBox "There Is Any Problem With Application!" & Chr(13) & _
    "Please Contact Us (PT. Solusi Sentral Data - (031) 5616465)", vbInformation, headerMSG
Exit Sub
End Sub

Private Sub newData()
    With frm_trans_att_man
        .DTPicker1.Value = Format(DTPicker1.Value, "yyyy-MM-dd")
        .isijamkerja
        .TDBCombo_company.Text = TDBCombo_company.Text
        .txt_company_name.Text = txt_company_name.Text
'        .txtkdshift.Text = TDBCombo1.Text
        .TDBCombo2.Text = TDBCombo1.Text
        .txtnmshift.Text = txtnmshift.Text
'        .TDBCombo2.Enabled = False
'        .txtnmshift.Enabled = False
        .cmbdep.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
        .txtnmDept.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
        .Show vbModal
    End With
End Sub

Private Sub Form_Resize()
    Frame1.Width = Me.Width - 500
    Frame1.Height = Me.Height - 1700
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

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode_rs(rscompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub vbButton1_Click()
    If LynxGrid2.Rows = 0 Then
        MsgBox "Invalid Departement Code." & Chr(13) & "Please Check Your Transaction Again...", vbInformation, "Error"
        Exit Sub
    End If
    If TDBCombo1.Text = "" Or TDBCombo1.Text = "ALL" Then
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
On Error GoTo err
    tanya = MsgBox("Deleted Data Can Not be Undo..!!", vbCritical + vbYesNo, "Warning")
    If tanya = vbYes Then
        CnG.Execute "DELETE FROM h_attendance where employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 0) & "' " _
            & "and att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
        MsgBox "Delete Data Sucsess..."
        
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute strsql
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        isiGridAbsen
    End If
'     and shift_code = '" & TDBCombo1.Text & "'
    Exit Sub

err:
MsgBox "There Is Any Problem With Application!" & Chr(13) & _
    "Please Contact Us (PT. Solusi Sentral Data - (031) 5616465)", vbInformation, headerMSG
Exit Sub
End Sub

Private Sub vbButton4_Click()
    isiGridAbsen
End Sub
