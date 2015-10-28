VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_list_overtime 
   Caption         =   "Manual Attendance"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17280
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   17280
   WindowState     =   2  'Maximized
   Begin prj_hardys.vbButton vbButton3 
      Height          =   345
      Left            =   15240
      TabIndex        =   13
      Top             =   600
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
      MICON           =   "frm_list_over_time.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_hardys.vbButton vbButton2 
      Height          =   345
      Left            =   13560
      TabIndex        =   12
      Top             =   600
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
      MICON           =   "frm_list_over_time.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_hardys.vbButton vbButton1 
      Height          =   345
      Left            =   12120
      TabIndex        =   11
      Top             =   600
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
      MICON           =   "frm_list_over_time.frx":0038
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
      TabIndex        =   8
      Top             =   600
      Width           =   2325
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1410
      OleObjectBlob   =   "frm_list_over_time.frx":0054
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
      TabIndex        =   4
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
      Format          =   108527617
      CurrentDate     =   40794
   End
   Begin VB.Frame Frame1 
      Height          =   9315
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   20025
      Begin prj_hardys.LynxGrid LynxGrid2 
         Height          =   7425
         Left            =   120
         TabIndex        =   10
         Top             =   210
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   13097
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
      Begin prj_hardys.LynxGrid LynxGrid1 
         Height          =   7425
         Left            =   3540
         TabIndex        =   9
         Top             =   210
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   13097
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
      Left            =   8370
      OleObjectBlob   =   "frm_list_over_time.frx":2012
      TabIndex        =   3
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Shift"
      Height          =   195
      Left            =   7710
      TabIndex        =   7
      Top             =   660
      Width           =   585
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
Attribute VB_Name = "frm_list_overtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim strSQL As String

Private Sub cmbshift_Change()
    LynxGrid1.Clear
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
    
    Call isiShift
    Call isiCompany
'    Call isiAbsentStatus
End Sub

Private Sub isiShift()
Dim rsshift As New ADODB.Recordset
    strSQL = "select shift_code,shift_name " _
            & "from m_shift"
    rsshift.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo1.RowSource = rsshift
End Sub

'Private Sub isiAbsentStatus()
'Dim rsabs As New ADODB.Recordset
'    strSQL = "select absent_code,absent_name " _
'            & "from m_absent_status"
'    rsabs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'    Set TDBCombo2.RowSource = rsabs
'    TDBCombo2.BoundText = "0"
'    txtabsentname.Text = "PRESENT"
'End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "NIK", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Employee Name", 3500, , , , , , , , True
      .AddColumn "Div. Code", 1000, , , , , , , , False
      .AddColumn "Division", 2000, , , , , , , , , True
      .AddColumn "Title Code", 1000, , , , , , , , False
      .AddColumn "Title", 2000, , , , , , , , , True
      .AddColumn "Cek IN", 1200, lgAlignCenterCenter, lgDate, "hh:ss", , , , , True
      .AddColumn "Break OUT", 1200, lgAlignCenterCenter, lgDate, "hh:ss", , , , , True
      .AddColumn "Break IN", 1200, lgAlignCenterCenter, lgDate, "hh:ss", , , , , True
      .AddColumn "Cek OUT", 1200, lgAlignCenterCenter, lgDate, "hh:ss", , , , , True
      .AddColumn "Entry Date", 1200, lgAlignCenterCenter, lgDate, "dd-MM-yyyy", , , , , True
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
    
    strSQL = "select * from m_company order by company_code"
    rscompany.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo_company.RowSource = rscompany
    'rs.Close

End Sub

Private Sub isiGridDep()
    LynxGrid2.Redraw = False
    LynxGrid2.Clear
    strSQL = "select department_code,department_name " _
            & "from m_department where company_code = '" & TDBCombo_company.Text & "'"
    rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
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
    Dim kdshift As String
    
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    strSQL = "select a.employee_code,b.employee_name," _
                & "b.division_code,c.division_name,b.title_code,d.title_name," _
                & "a.time_in,a.time_out_break,a.time_in_break,a.time_out,a.entry_date " _
        & "from m_employee b LEFT JOIN h_attendance a on b.employee_code = a.employee_code " _
        & "join m_division c on b.division_code = c.division_code " _
        & "join m_title d on b.title_code = d.title_code " _
        & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " _
        & "AND date(a.att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " _
        & "AND a.shift_code = '" & TDBCombo1.Text & "' " _
        & "AND b.department_code = '" & LynxGrid2.CellText(LynxGrid2.Row, 0) & "'"
    rs.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!employee_code & vbTab & rs!employee_name _
                & vbTab & rs!division_code & vbTab & rs!division_name _
                & vbTab & rs!title_code & vbTab & rs!title_name _
                & vbTab & rs!time_in & vbTab & rs!time_out_break & vbTab & rs!time_in_break & vbTab & rs!time_out _
                & vbTab & rs!entry_date
            rs.MoveNext
        Wend
    End If
    rs.Close
    LynxGrid1.Redraw = True
End Sub

Private Sub showData()
    If LynxGrid1.Rows > 0 Then
        With frm_trans_att_man
            .DTPicker1.Value = DTPicker1.Value
            .TDBCombo_company.Text = TDBCombo_company.Text
            .txt_company_name.Text = txt_company_name.Text
            .txtkdshift.Text = TDBCombo1.Text
            .txtnmshift.Text = txtnmshift.Text
            .cmbdep.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            .txtnmDept.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            .txtkdkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 0)
            .txtnmkar.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
            .txtkddiv.Text = LynxGrid1.CellText(LynxGrid1.Row, 2)
            .txtdivision.Text = LynxGrid1.CellText(LynxGrid1.Row, 3)
            .txtkdtitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 4)
            .txttitle.Text = LynxGrid1.CellText(LynxGrid1.Row, 5)
            .ttin.Value = Format(LynxGrid1.CellText(LynxGrid1.Row, 6), "hh:ss")
            .ttout.Value = Format(LynxGrid1.CellText(LynxGrid1.Row, 9), "hh:ss")
            .txtentry.Text = LynxGrid1.CellText(LynxGrid1.Row, 6)
'            .TDBCombo1.BoundText = TDBCombo2.Text
'            .txtnmabsentstatus.Text = txtabsentname.Text
            
            .editTrans = True
            .Show vbModal
        End With
    End If
End Sub

Private Sub newData()
    With frm_trans_att_man
        .DTPicker1.Value = DTPicker1.Value
        .TDBCombo_company.Text = TDBCombo_company.Text
        .txt_company_name.Text = txt_company_name.Text
        .txtkdshift.Text = TDBCombo1.Text
        .txtnmshift.Text = txtnmshift.Text
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
    tanya = MsgBox("Deleted Data Can Not be Undo..!!", vbCritical + vbYesNo, "Warning")
    If tanya = vbYes Then
        CnG.Execute "DELETE FROM h_attendance where employee_code = '" & LynxGrid1.CellText(LynxGrid1.Row, 0) & "' " _
            & "and att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' and shift_code = '" & TDBCombo1.Text & "'"
        MsgBox "Delete Data Sucsess..."
        isiGridAbsen
    End If
End Sub
