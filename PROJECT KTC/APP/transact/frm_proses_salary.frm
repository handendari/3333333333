VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_proses_salary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Process Calculate Salary"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9045
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   7200
      Top             =   840
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   2790
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   13
      Top             =   390
      Width           =   2745
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   330
      TabIndex        =   12
      Top             =   2820
      Visible         =   0   'False
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin prj_fej_jkt.vbButton vbButton1 
      Height          =   465
      Left            =   510
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "Process"
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
      MICON           =   "frm_proses_salary.frx":0000
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
      Left            =   1530
      TabIndex        =   4
      Top             =   1350
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   92078083
      CurrentDate     =   40842
   End
   Begin VB.ComboBox cmbBulan 
      Height          =   315
      ItemData        =   "frm_proses_salary.frx":001C
      Left            =   1560
      List            =   "frm_proses_salary.frx":0044
      TabIndex        =   2
      Text            =   "1"
      Top             =   840
      Width           =   885
   End
   Begin VB.TextBox txtTahun 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2460
      TabIndex        =   3
      Text            =   "2011"
      Top             =   840
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   1350
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   92078083
      CurrentDate     =   40842
   End
   Begin prj_fej_jkt.vbButton vbButton2 
      Height          =   465
      Left            =   4110
      TabIndex        =   7
      Top             =   3240
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "CANCEL"
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
      MICON           =   "frm_proses_salary.frx":0078
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
      Height          =   465
      Left            =   5760
      TabIndex        =   8
      Top             =   3240
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "EXIT"
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
      MICON           =   "frm_proses_salary.frx":0094
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   2820
      Top             =   360
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1560
      OleObjectBlob   =   "frm_proses_salary.frx":00B0
      TabIndex        =   1
      Top             =   390
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "To "
      Height          =   195
      Left            =   3150
      TabIndex        =   11
      Top             =   1410
      Width           =   285
   End
   Begin VB.Label Label3 
      Caption         =   "Branch Office :"
      Height          =   195
      Left            =   390
      TabIndex        =   10
      Top             =   450
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Periode :"
      Height          =   195
      Left            =   840
      TabIndex        =   9
      Top             =   1440
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Salary Month :"
      Height          =   225
      Left            =   450
      TabIndex        =   0
      Top             =   900
      Width           =   1065
   End
End
Attribute VB_Name = "frm_proses_salary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim strsql As String
Dim v_tgl_awal, v_tgl_akhir As Date

Private Sub cmbBulan_Click()
DTPicker1.Value = txtTahun.Text & "-" & cmbBulan.Text & "-01"
DTPicker2.Value = txtTahun.Text & "-" & cmbBulan.Text & "-" & getEndDay(Val(cmbBulan.Text), Val(txtTahun.Text))
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Private Sub txtTahun_Validate(Cancel As Boolean)
DTPicker1.Value = txtTahun.Text & "-" & cmbBulan.Text & "-01"
DTPicker2.Value = txtTahun.Text & "-" & cmbBulan.Text & "-" & getEndDay(Val(cmbBulan.Text), Val(txtTahun.Text))
End Sub

Private Sub Form_Load()
    Adodc_company.ConnectionString = strConn
    
'    DTPicker1.Value = Now()
'    DTPicker2.Value = Now()

    txtTahun.Text = Format(Now(), "yyyy")
    cmbBulan.Text = Format(Now(), "MM")
    
    DTPicker1.Value = txtTahun.Text & "-" & cmbBulan.Text & "-01"
    DTPicker2.Value = txtTahun.Text & "-" & cmbBulan.Text & "-" & getEndDay(Val(cmbBulan.Text), Val(txtTahun.Text))

    Call load_data_company
    
    timer1.Enabled = True
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If
End Sub

Public Sub load_data_company()
    Adodc_company.RecordSource = "select * from m_company order by company_code"
    Adodc_company.Refresh

    TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub vbButton1_Click()
Dim bulan As String
Dim tgl As String
Dim v_tgl_akhir As Date
Dim v_tgl_mc As Date
Dim v_end_mc As Date
Dim int_month As Integer
Dim int_year As Integer
Dim clsinsert As New clsInsert_h_salary
Dim bln_awal, bln_akhir, thn_awal, thn_akhir As String

bln_awal = Mid((Format(DTPicker1.Value, "yyyy-MM-dd")), 6, 2)
bln_akhir = Mid((Format(DTPicker2.Value, "yyyy-MM-dd")), 6, 2)
thn_awal = Year(Format(DTPicker1.Value, "yyyy-MM-dd"))
thn_akhir = Year(Format(DTPicker2.Value, "yyyy-MM-dd"))

If (bln_awal <> cmbBulan.Text Or thn_awal <> txtTahun.Text) Or (bln_akhir <> cmbBulan.Text Or thn_akhir <> txtTahun.Text) Then
    MsgBox "Periode Did Not Match With Month Or Year!" & Chr(13) & "Please Check Your Input Periode.." _
        , vbExclamation, "Error!"
    Exit Sub
End If

bulan = txtTahun.Text + "-" + cmbBulan.Text
tgl = txtTahun.Text + "-" + cmbBulan.Text + "-01"

    ProgressBar1.Visible = True
    
    strsql = "Select LAST_DAY('" & tgl & "') tgl_akhir, employee_code, employee_name," & _
            "marital_status, number_of_children, sex," & _
            "end_working, start_mc, CAST(IFNULL(end_mc,LAST_DAY('" & tgl & "')) as DATE) end_mc, flag_active " & _
        "FROM m_employee " & _
        "WHERE company_code = '" & TDBCombo_company.Text & "' AND flag_active > 0"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    v_tgl_akhir = rs!tgl_akhir
    v_end_mc = IIf(IsNull(rs!end_mc), "00:00:00", rs!end_mc)
           
    If rs.RecordCount > 0 Then
        ProgressBar1.Max = rs.RecordCount
        ProgressBar1.Value = 0
        
        CnG.BeginTrans
        
        strsql = "DELETE FROM h_salary_new WHERE month = '" & bulan & "' " _
            & "AND company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute strsql
        
        rs.MoveFirst
        While Not rs.EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            
            v_tgl_mc = IIf(IsNull(rs!start_mc), 0, rs!start_mc)
            int_month = Month(v_tgl_akhir)
            int_year = Year(v_tgl_akhir)
            
            If rs!flag_active = "2" Then
            
                strsql = "DELETE FROM h_attendance " & _
                        "WHERE employee_code = '" & rs!employee_code & "' AND DATE(att_date) >= '" & Format(v_tgl_mc, "yyyy-MM-dd") & "' " & _
                        "AND month(att_date) = '" & cmbBulan.Text & "' AND year(att_date) = '" & txtTahun.Text & "'"
                CnG.Execute strsql
                
                    'rs.MoveFirst
                If Format(v_tgl_mc, "yyyy-MM") = Format(v_tgl_akhir, "yyyy-MM") Then
                    v_tgl_mc = DateValue(v_tgl_mc)
                    While v_tgl_mc <= v_tgl_akhir
'                        v_tgl_mc = v_tgl_mc + 1
                        strsql = "INSERT INTO h_attendance (employee_code, att_date," & _
                            "shift_number, shift_code, start_time," & _
                            "end_time," & _
                            "flag_present,description, entry_date) " & _
                        "VALUES " & _
                            "('" & rs!employee_code & "','" & Format(v_tgl_mc, "yyyy-MM-dd 00:00:00") & "'," & _
                            "1, 'MC','" & Format(v_tgl_mc, "yyyy-MM-dd 08:00:00") & "'," & _
                            "'" & Format(v_tgl_mc, "yyyy-MM-dd 17:00:00") & "'," & _
                            "0,'MC', Now())"
                        
                        CnG.Execute strsql
                        'rs.MoveNext
                        v_tgl_mc = v_tgl_mc + 1
                    Wend
                Else
                    Dim i As Integer
                    For i = 1 To Day(v_tgl_akhir)
'                        i = i + 1
'                        v_tgl_mc = Format(v_tgl_akhir, "yyyy-MM") + "-" + i
                        v_tgl_mc = DateSerial(int_year, int_month, i)
                        strsql = "INSERT INTO h_attendance (employee_code, att_date," & _
                            "shift_number, shift_code, start_time," & _
                            "end_time," & _
                            "flag_present,description, entry_date) " & _
                        "VALUES " & _
                            "('" & rs!employee_code & "','" & Format(v_tgl_mc, "yyyy-MM-dd 00:00:00") & "'," & _
                            "1, 'MC','" & Format(v_tgl_mc, "yyyy-MM-dd 08:00:00") & "'," & _
                            "'" & Format(v_tgl_mc, "yyyy-MM-dd 17:00:00") & "'," & _
                            "0,'MC', Now())"
    
                        CnG.Execute strsql
                    Next
                End If
            End If
            
            Call clsinsert.Insert_h_salary(rs!employee_code, rs!sex, bulan, Format(DTPicker1.Value, "yyyy-MM-dd"), _
                Format(DTPicker2.Value, "yyyy-MM-dd"), rs!marital_status, IIf(IsNull(rs!number_of_children), 0, rs!number_of_children), _
                IIf(IsNull(Format(rs!start_mc, "yyyyMM")), "0", Format(rs!start_mc, "yyyyMM")), _
                rs!flag_active, TDBCombo_company.Text)
            
            'Update Employee Loan +++++++++++++++++++++++++++++++++++++
            strsql = "UPDATE td_loan SET flag_paid = 1 " _
                & "Where employee_code = '" & rs!employee_code & "' " _
                & "AND Month(installment_month) = '" & cmbBulan.Text & "' " _
                & "AND Year(installment_month) = '" & txtTahun.Text & "'"
            CnG.Execute (strsql)
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            'Update Temp Salary Proses ++++++++++++++++++++++++++++++++
            strsql = "UPDATE temp_sal_proses set salary_proses = 1 where company_code = '" & TDBCombo_company.Text & "'"
            CnG.Execute (strsql)
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            rs.MoveNext
        Wend
            
            'Update m_day +++++++++++++++++++++++++++++++++++++++++++++
            v_tgl_awal = Format(DTPicker1.Value, "yyyy-MM-dd")
            v_tgl_akhir = Format(DTPicker2.Value, "yyyy-MM-dd")
            
            v_tgl_awal = DateValue(v_tgl_awal)
            v_tgl_akhir = DateValue(v_tgl_akhir)
            
'            strsql = "delete from m_day where dt between '" & v_tgl_awal & "' and '" & v_tgl_akhir & "'"
'            CnG.Execute strsql
        
            While v_tgl_awal <= v_tgl_akhir
                DoEvents
                
                strsql = "delete from m_day where dt = '" & Format(v_tgl_awal, "yyyy-MM-dd") & "'"
                CnG.Execute strsql
                
                strsql = "INSERT INTO m_day (dt) VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "')"
                CnG.Execute strsql
                
                v_tgl_awal = v_tgl_awal + 1
            Wend
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        CnG.CommitTrans
    End If
    rs.Close
    
    CnG.Execute "call spg_leave_periode2 ('" & Format(DTPicker2.Value, "yyyy-MM-dd") & "')"
ProgressBar1.Visible = False
End Sub

Private Sub vbButton2_Click()
    Call load_data_company
    
    txtTahun.Text = Format(Now(), "yyyy")
    cmbBulan.Text = Format(Now(), "MM")
    
    TDBCombo_company.Text = ""
    txt_company_name.Text = ""
    
    DTPicker1.Value = Now()
    DTPicker2.Value = Now()
End Sub

Private Sub vbButton3_Click()
    Unload Me
End Sub
