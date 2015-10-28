VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_reproccess_download 
   Caption         =   "RE PROCCESS"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7740
   Icon            =   "frm_reproccess_download.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4215
   ScaleWidth      =   7740
   StartUpPosition =   1  'CenterOwner
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   1965
      Left            =   3330
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3466
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
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   210
      Top             =   3360
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2970
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   12
      Top             =   180
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Height          =   2625
      Left            =   210
      TabIndex        =   0
      Top             =   690
      Width           =   7275
      Begin VB.ComboBox cbo_department 
         Height          =   315
         ItemData        =   "frm_reproccess_download.frx":058A
         Left            =   1290
         List            =   "frm_reproccess_download.frx":0594
         TabIndex        =   10
         Text            =   "..."
         Top             =   630
         Width           =   1695
      End
      Begin VB.Frame fra_department 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   3120
         TabIndex        =   7
         Top             =   540
         Visible         =   0   'False
         Width           =   4005
         Begin VB.TextBox txt_department_name 
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
            Left            =   1380
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   90
            Width           =   2595
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_department 
            Height          =   375
            Left            =   0
            OleObjectBlob   =   "frm_reproccess_download.frx":05A7
            TabIndex        =   9
            Top             =   90
            Width           =   1365
         End
      End
      Begin VB.ComboBox cbo_employee 
         Height          =   315
         ItemData        =   "frm_reproccess_download.frx":2510
         Left            =   1290
         List            =   "frm_reproccess_download.frx":251A
         TabIndex        =   1
         Text            =   "..."
         Top             =   1020
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker_periode_from 
         Height          =   300
         Left            =   1290
         TabIndex        =   2
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   122159107
         CurrentDate     =   39278
      End
      Begin MSComCtl2.DTPicker DTPicker_periode_to 
         Height          =   300
         Left            =   3540
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   122159107
         CurrentDate     =   39278
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   2280
         Visible         =   0   'False
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Frame fra_employee 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   615
         Left            =   3120
         TabIndex        =   18
         Top             =   780
         Width           =   4065
         Begin VB.TextBox txt_employee_code 
            Height          =   315
            Left            =   3690
            TabIndex        =   23
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txt_employee_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   20
            Top             =   240
            Width           =   2145
         End
         Begin VB.TextBox txt_nik 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            Height          =   315
            Left            =   0
            MaxLength       =   50
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
         Begin prj_tpc.vbButton cmdBrowse 
            Height          =   315
            Left            =   1380
            TabIndex        =   21
            Top             =   240
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            BTYPE           =   14
            TX              =   "..."
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
            MICON           =   "frm_reproccess_download.frx":252D
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
      Begin MSComCtl2.DTPicker DTPicker_Periode 
         Height          =   315
         Left            =   1290
         TabIndex        =   24
         Top             =   1410
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM-yyyy"
         Format          =   122159107
         CurrentDate     =   40794
      End
      Begin VB.Label Label5 
         Caption         =   "Periode"
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Department"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Date"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Employee"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1050
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   3030
         TabIndex        =   4
         Top             =   1830
         Width           =   465
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1170
      OleObjectBlob   =   "frm_reproccess_download.frx":2549
      TabIndex        =   13
      Top             =   180
      Width           =   1695
   End
   Begin prj_tpc.vbButton cmd_close 
      Height          =   705
      Left            =   6540
      TabIndex        =   16
      Top             =   3390
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1244
      BTYPE           =   14
      TX              =   "&Exit"
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
      MICON           =   "frm_reproccess_download.frx":44AF
      PICN            =   "frm_reproccess_download.frx":44CB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmd_use 
      Height          =   705
      Left            =   5520
      TabIndex        =   17
      Top             =   3390
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1244
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
      MICON           =   "frm_reproccess_download.frx":555D
      PICN            =   "frm_reproccess_download.frx":5579
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   210
      TabIndex        =   14
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frm_reproccess_download"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset

Private Sub cbo_department_Click()
    If cbo_department.ListIndex = 0 Then
        fra_department.Visible = False
        TDBCombo_department.Text = ""
        txt_department_name.Text = ""
        TDBCombo_department.RowSource = Nothing
    Else
        fra_department.Visible = True
        Call load_data_department
    End If
End Sub

Private Sub cbo_employee_Click()
    If cbo_employee.ListIndex = 0 Then
        fra_employee.Visible = False
        txt_nik.Text = ""
        txt_employee_name.Text = ""
        txt_employee_code = ""
    Else
        fra_employee.Visible = True
    End If
End Sub

Private Sub cmd_close_Click()
    Unload Me
End Sub

Private Sub cmd_use_Click()
Dim vEnrollNumber As String
On Error Resume Next
    SQL = "DELETE FROM h_attendance_reproccess"
    CnG.Execute SQL
    
    SQL = "DELETE FROM h_log_attendance_recover"
    CnG.Execute SQL
    
    SQL = "DELETE FROM h_log_attendance_reproccess"
    CnG.Execute SQL
       
    CnG.BeginTrans
    
    If cbo_employee.ListIndex = 0 And cbo_department.ListIndex = 0 Then
        SQL = "INSERT INTO h_attendance_reproccess " & _
              "SELECT * FROM h_attendance " & _
              "WHERE IFNULL(flag_manual,0) = 1 " & _
                  "AND (DATE(att_date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " & _
                      "AND '" & Format(DateAdd("d", 1, DTPicker_periode_to.Value), "yyyy-MM-dd") & "')"
        CnG.Execute SQL


        SQL = "INSERT INTO h_log_attendance_recover(att_date, ip_address, enrollnumber, employee_code, verifymode, " _
                            & "flag_io, flag_attendance, ref_date, entry_date) " _
                    & "SELECT att_date, ip_address, enrollnumber, employee_code, verifymode, flag_io, " _
                            & "flag_attendance , ref_date, entry_date " _
                    & "FROM h_log_attendance " _
                    & "WHERE (date(att_date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " _
                        & "AND '" & Format(DateAdd("d", 1, DTPicker_periode_to.Value), "yyyy-MM-dd") & "') " _
                    & "ORDER BY att_date, flag_io"
        CnG.Execute SQL
        
        SQL = "DELETE FROM h_attendance " _
                & "WHERE (date(att_date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " _
                    & "AND '" & Format(DateAdd("d", 1, DTPicker_periode_to.Value), "yyyy-MM-dd") & "')"
        CnG.Execute SQL
    ElseIf cbo_employee.ListIndex = 0 And cbo_department.ListIndex = 1 Then
        SQL = "INSERT INTO h_attendance_reproccess " & _
              "SELECT a.* FROM h_attendance a JOIN m_employee b ON a.employee_code = b.employee_code " & _
              "WHERE IFNULL(a.flag_manual,0) = 1 " & _
                  "AND (DATE(a.att_date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " & _
                      "AND '" & Format(DateAdd("d", 1, DTPicker_periode_to.Value), "yyyy-MM-dd") & "') " & _
                      "AND b.department_code = '" & TDBCombo_department.Text & "'"
        CnG.Execute SQL
        
        SQL = "INSERT INTO h_log_attendance_recover(att_date, ip_address, enrollnumber, employee_code, verifymode, " _
                            & "flag_io, flag_attendance, ref_date, entry_date) " _
                    & "SELECT a.att_date, a.ip_address, a.enrollnumber, a.employee_code, a.verifymode, a.flag_io, " _
                            & "a.flag_attendance , a.ref_date, a.entry_date " _
                    & "FROM h_log_attendance a join m_employee b on a.enrollnumber = (b.nik+0) " _
                    & "WHERE (date(att_date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " _
                        & "AND '" & Format(DateAdd("d", 1, DTPicker_periode_to.Value), "yyyy-MM-dd") & "') AND b.department_code = '" & TDBCombo_department.Text & "' " _
                    & "ORDER BY att_date, flag_io"
        CnG.Execute SQL
        
        SQL = "DELETE a FROM h_attendance a join m_employee b on a.employee_code = b.employee_code " _
                & "WHERE (date(att_date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " _
                    & "AND '" & Format(DateAdd("d", 1, DTPicker_periode_to.Value), "yyyy-MM-dd") & "') " _
                    & "AND b.department_code = '" & TDBCombo_department.Text & "'"
        CnG.Execute SQL
    Else
        SQL = "INSERT INTO h_attendance_reproccess " & _
              "SELECT * FROM h_attendance " & _
              "WHERE IFNULL(flag_manual,0) = 1 " & _
                  "AND (DATE(att_date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " & _
                      "AND '" & Format(DateAdd("d", 1, DTPicker_periode_to.Value), "yyyy-MM-dd") & "') " & _
                      "AND employee_code = '" & txt_employee_code.Text & "'"
        CnG.Execute SQL
        
        SQL = "SELECT enrollnumber FROM m_enroll_link WHERE employee_code = '" & txt_employee_code.Text & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            While Not rs.EOF
                SQL = "INSERT INTO h_log_attendance_recover(att_date, ip_address, enrollnumber, employee_code, verifymode, " _
                                & "flag_io, flag_attendance, ref_date, entry_date) " _
                        & "SELECT att_date, ip_address, enrollnumber, employee_code, verifymode, flag_io, " _
                                & "flag_attendance , ref_date, entry_date " _
                        & "FROM h_log_attendance " _
                        & "WHERE (date(att_date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " _
                            & "AND '" & Format(DateAdd("d", 1, DTPicker_periode_to.Value), "yyyy-MM-dd") & "') " _
                            & "AND enrollnumber = '" & Val(rs!enrollnumber) & "' " _
                        & "ORDER BY att_date, flag_io"
                CnG.Execute SQL
            rs.MoveNext
            Wend
        End If
        rs.Close
        
        SQL = "DELETE FROM h_attendance " _
                & "WHERE (date(att_date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " _
                    & "AND '" & Format(DateAdd("d", 1, DTPicker_periode_to.Value), "yyyy-MM-dd") & "') " _
                    & "AND employee_code = '" & txt_employee_code.Text & "'"
        CnG.Execute SQL
    End If
    
    If rs.State Then rs.Close
    
    SQL = "SELECT * from h_log_attendance_recover ORDER BY att_date, flag_io"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    ProgressBar1.Visible = True
    ProgressBar1.Value = 0
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            ProgressBar1.Max = rs.RecordCount
            ProgressBar1.Value = ProgressBar1.Value + 1
            
            SQL = "INSERT INTO h_log_attendance_reproccess(att_date, ip_address, enrollnumber, employee_code, verifymode, " _
                            & "flag_io, flag_attendance, ref_date, entry_date) VALUES (" _
                    & "'" & Format(rs!att_date, "yyyy-MM-dd hh:nn:ss") & "', '" & rs!ip_address & "', '" & IIf(IsNull(rs!enrollnumber), 0, rs!enrollnumber) & "', " _
                    & "'" & rs!employee_code & "', '" & IIf(IsNull(rs!verifymode), 0, rs!verifymode) & "', '" & rs!flag_io & "', " _
                    & "'" & IIf(IsNull(rs!flag_attendance), 0, rs!flag_attendance) & "' , '" & Format(IIf(IsNull(rs!ref_date), 0, rs!ref_date), "yyyy-MM-dd hh:nn:ss") & "', '" & Format(IIf(IsNull(rs!entry_date), 0, rs!entry_date), "yyyy-MM-dd hh:nn:ss") & "')"
            CnG.Execute SQL
        rs.MoveNext
        Wend
    End If
    rs.Close
    
    SQL = "DELETE a FROM h_attendance a JOIN h_attendance_reproccess b ON a.employee_code = b.employee_code AND a.att_date = b.att_date"
    CnG.Execute SQL
    
    SQL = "INSERT INTO h_attendance " & _
          "SELECT * FROM h_attendance_reproccess"
    CnG.Execute SQL
    
    Screen.MousePointer = vbNormal
    
    CnG.CommitTrans
    
    ProgressBar1.Visible = False
    MsgBox "Proccess Successfully!", vbInformation, headerMSG
End Sub

Private Sub DTPicker_Periode_Change()
    Call getPeriode(DTPicker_Periode.Value, DTPicker_periode_from, DTPicker_periode_to)
End Sub

Private Sub Form_Load()
    Call load_data_company
    Call createGridKar
    
    cbo_department.ListIndex = 0
    cbo_employee.ListIndex = 0
    
    DTPicker_Periode.Value = Now
    DTPicker_periode_from.Value = Now
    DTPicker_periode_to.Value = Now
    
    timer1.Enabled = True
End Sub
Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub load_data_department()
    If rsDept.State Then rsDept.Close
    SQL = "select * from m_department where company_code = '" & TDBCombo_company.Text & "' " _
            & "order by department_code"
    rsDept.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_department.RowSource = rsDept
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_reproccess_download = Nothing
End Sub

Private Sub TDBCombo_department_ItemChange()
    If TDBCombo_department.ApproxCount > 0 Then
        TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
        txt_department_name = TDBCombo_department.Columns("department_name").Value
    End If
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub createGridKar()
   With LynxGrid2
      .AddColumn "Employee Code", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "employee_code", 2000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear
        
        vParam = IIf(DEPARTMENT_CODE <> "" And DIVISION_CODE = "", "a.department_code = '" & DEPARTMENT_CODE & "'", IIf(DEPARTMENT_CODE = "" And DIVISION_CODE = "", "a.company_code = '" & COMPANY_CODE & "'", "a.department_code = '" & DEPARTMENT_CODE & "' AND a.division_code = '" & DIVISION_CODE & "'"))
        
        If LOGIN_LEVEL = 100 Then
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_nik.Text & "%')"
        Else
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND " & vParam & " " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_nik.Text & "%') " & _
                        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
        End If
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid2.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME & vbTab & rs!employee_code
                rs.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                txt_employee_code.Text = rs!employee_code
                txt_employee_name.Text = rs!EMPLOYEE_NAME
                txt_nik.Text = rs!nik
'                TDBCombo1.SetFocus
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else
            
        End If
        rs.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txt_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txt_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txt_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub LynxGrid2_DblClick()
    isiGridKar (2)
End Sub

Private Sub LynxGrid2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        LynxGrid2.Visible = False
    End If
    If KeyAscii = 13 Then
        isiGridKar (2)
    End If
End Sub

Private Sub LynxGrid2_LostFocus()
    LynxGrid2.Visible = False
End Sub

Private Sub txt_nik_Change()
    If txt_nik.Text = "" Then
        txt_employee_code.Text = ""
        txt_employee_name.Text = ""
    End If
End Sub

Private Sub txt_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmdBrowse_Click()
    isiGridKar (1)
End Sub

