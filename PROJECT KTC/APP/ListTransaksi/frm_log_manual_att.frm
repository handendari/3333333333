VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_log_manual_att 
   Caption         =   "Log Manual Attendance"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13965
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   13965
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   7500
      Top             =   180
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1410
      OleObjectBlob   =   "frm_log_manual_att.frx":0000
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   990
      TabIndex        =   1
      Top             =   210
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      Format          =   93061121
      CurrentDate     =   40794
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
         TabIndex        =   7
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
         TabIndex        =   2
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
   Begin VB.Label Label1 
      Caption         =   "Date :"
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   270
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Branch Office :"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   660
      Width           =   1065
   End
End
Attribute VB_Name = "frm_log_manual_att"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim strsql As String
Dim rscompany As New ADODB.Recordset

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

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "NIK", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Employee Name", 3500, , , , , , , , True
      .AddColumn "Div. Code", 1000, , , , , , , , False
      .AddColumn "Division", 2000, , , , , , , , , True
      .AddColumn "Title Code", 1000, , , , , , , , False
      .AddColumn "Title", 2000, , , , , , , , , True
      .AddColumn "Shift Code", 1000, , , , , , , , , True
      .AddColumn "Shift Name", 1000, , , , , , , , False
      .AddColumn "Cek IN", 1200, lgAlignCenterCenter, lgDate, "hh:ss", , , , , True
      .AddColumn "Break OUT", 1200, lgAlignCenterCenter, lgDate, "hh:ss", , , , , True
      .AddColumn "Break IN", 1200, lgAlignCenterCenter, lgDate, "hh:ss", , , , , True
      .AddColumn "Cek OUT", 1200, lgAlignCenterCenter, lgDate, "hh:ss", , , , , True
      .AddColumn "Entry Date", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , True
      .AddColumn "Absent", 1000, , , , , , , , False
      .AddColumn "Absent_name", 1000, , , , , , , , False
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

Public Sub isiGridAbsen()
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    strsql = "select b.employee_code,b.employee_name,a.absent_status,a.absent_name," _
                & "b.division_code,c.division_name,b.title_code,d.title_name,a.shift_code,a.shift_name," _
                & "a.time_in,a.time_out_break,a.time_in_break,a.time_out,a.entry_date " _
        & "from m_employee b LEFT JOIN " _
            & "(SELECT employee_code,xx.shift_code,xx.absent_status,xy.absent_name,yy.shift_name,time_in,time_out_break,time_in_break,time_out,entry_date " & _
                "FROM h_attendance xx left join m_shift yy on xx.shift_code = yy.shift_code " _
                & "join m_absent_status xy on xx.absent_status = xy.absent_code " _
                & "WHERE date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "') a " _
            & "on b.employee_code = a.employee_code " _
        & "join m_division c on b.division_code = c.division_code and b.company_code = c.company_code " _
        & "join m_title d on b.title_code = d.title_code " _
        & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " _
        & "AND b.department_code = '" & Replace(LynxGrid2.CellText(LynxGrid2.Row, 0), "'", "''") & "'"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!employee_code & vbTab & rs!employee_name _
                & vbTab & rs!division_code & vbTab & rs!division_name _
                & vbTab & rs!title_code & vbTab & rs!title_name & vbTab & rs!shift_code & vbTab & rs!shift_name _
                & vbTab & rs!time_in & vbTab & rs!time_out_break & vbTab & rs!time_in_break & vbTab & rs!time_out _
                & vbTab & rs!entry_date & vbTab & rs!absent_status & vbTab & rs!absent_name
            rs.MoveNext
        Wend
    End If
    rs.Close
    LynxGrid1.Redraw = True
End Sub

Private Sub showData()
    If LynxGrid1.Rows > 0 Then
        frm_trans_att_man.DTPicker1.Value = DTPicker1.Value
        frm_trans_att_man.TDBCombo_company.Text = TDBCombo_company.Text
        frm_trans_att_man.txt_company_name.Text = txt_company_name.Text
        frm_trans_att_man.txtkdshift.Text = LynxGrid1.CellText(LynxGrid1.Row, 6)
        frm_trans_att_man.txtnmshift.Text = LynxGrid1.CellText(LynxGrid1.Row, 7)
        frm_trans_att_man.cmbdep.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
        frm_trans_att_man.txtnmDept.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
        frm_trans_att_man.txtkdkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 0)
        frm_trans_att_man.txtnmkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
        frm_trans_att_man.txtkddiv.Text = LynxGrid1.CellText(LynxGrid1.Row, 2)
        frm_trans_att_man.txtdivision.Text = LynxGrid1.CellText(LynxGrid1.Row, 3)
        frm_trans_att_man.txtkdtitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 4)
        frm_trans_att_man.txttitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 5)
        frm_trans_att_man.ttin.Value = Format(LynxGrid1.CellText(LynxGrid1.Row, 8), "hh:ss")
        frm_trans_att_man.ttout.Value = Format(LynxGrid1.CellText(LynxGrid1.Row, 11), "hh:ss")
        frm_trans_att_man.txtentry.Text = LynxGrid1.CellText(LynxGrid1.Row, 8)
        frm_trans_att_man.TDBCombo1.Text = LynxGrid1.CellText(LynxGrid1.Row, 13)
        frm_trans_att_man.txtnmabsentstatus.Text = LynxGrid1.CellText(LynxGrid1.Row, 14)
        frm_trans_att_man.editTrans = True
        frm_trans_att_man.Show vbModal
    End If
End Sub

Private Sub newData()
    frm_trans_att_man.DTPicker1.Value = DTPicker1.Value
    frm_trans_att_man.TDBCombo_company.Text = TDBCombo_company.Text
    frm_trans_att_man.txt_company_name.Text = txt_company_name.Text
    frm_trans_att_man.txtkdshift.Text = TDBCombo1.Text
    frm_trans_att_man.txtnmshift.Text = txtnmshift.Text
    frm_trans_att_man.cmbdep.Text = LynxGrid2.CellText(LynxGrid2.Row, 0) & " " & LynxGrid2.CellText(LynxGrid2.Row, 1)
    frm_trans_att_man.Show vbModal
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

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call set_company_mode_rs(rscompany, TDBCombo_company, txt_company_name)
End Sub
