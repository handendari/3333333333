VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_trans_import_log_attendance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IMPORT LOG ATTENDANCE DATA FROM MACHINE"
   ClientHeight    =   7815
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   14700
   Icon            =   "frm_trans_log_import_att.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   4500
      TabIndex        =   4
      Top             =   1530
      Width           =   5655
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   270
         TabIndex        =   5
         Top             =   2100
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   300
         TabIndex        =   7
         Top             =   1740
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Processing, Please Wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   6
         Top             =   300
         Width           =   2730
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7500
      Top             =   7200
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   582
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      Picture         =   "frm_trans_log_import_att.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmd_search 
      Caption         =   "&Browse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "frm_trans_log_import_att.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      Picture         =   "frm_trans_log_import_att.frx":0B20
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmd_refresh 
      Cancel          =   -1  'True
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Picture         =   "frm_trans_log_import_att.frx":10AA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prj_genting.LynxGrid LynxGrid1 
      Height          =   6435
      Left            =   270
      TabIndex        =   8
      Top             =   270
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   11351
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
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   -1  'True
      AllowColumnResizing=   -1  'True
      ColumnSort      =   -1  'True
   End
End
Attribute VB_Name = "frm_trans_import_log_attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v_jmljam As Double
'Dim v_jam_kerja As Double, v_jml_lembur As Double
Dim v_15 As Double, v_2 As Double, v_3 As Double, v_4 As Double
Dim v_tot_overtime As Double
Dim v_start_time As String
Dim v_end_time As String
Dim v_tot_ot As Double
Dim v_interval As Double
Dim tgl_in As String, tgl_out As String
Dim abstatus As Integer
Dim v_flag_ot As Double
Dim v_meal As Double, v_trans As Double
Dim v_hari_libur As Integer
Dim v_nik_master As String, v_employee_name_master As String
Dim v_nik As String, v_employee_name As String
Dim v_tot_work As Double
Dim v_company_code As String

Private Sub fill_grid_excel_m(str_file_name As String)
Dim strWorksheet As String
Dim total_OT As Double

On Error GoTo Err
    Me.MousePointer = vbHourglass
    'DoEvents
    LynxGrid1.Clear
    
    strWorksheet = "attendance"
    'strWorksheet_m = "tm_bom": strWorksheet_d = "td_bom"
    
    Adodc1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & str_file_name & ";Extended Properties=Excel 8.0"
    
    Adodc1.RecordSource = "select * from [" & strWorksheet & "$]"
    Adodc1.Refresh
    
    With Adodc1.Recordset
        If .RecordCount > 0 Then
            .MoveFirst
            While Not .EOF
                
                If IsNull(.Fields(2)) = False And .Fields(1).Type = adDate Then
'                    If IsNull(.Fields(6)) = False And IsNull(.Fields(7)) = False Then
                        LynxGrid1.AddItem .Fields(1) & vbTab & .Fields(2) & vbTab & .Fields(3) & vbTab & .Fields(4) _
                        & vbTab & .Fields(5) & vbTab & .Fields(6) & vbTab & .Fields(7) & vbTab & .Fields(8) _
                        & vbTab & .Fields(12) & vbTab & .Fields(11) & vbTab & .Fields(13)
'                    End If
                End If
                .MoveNext
            Wend
            LynxGrid1.Redraw = True
        End If
    End With
    Me.MousePointer = vbNormal
    Exit Sub
    
Err:
    MsgBox Err.Description, vbExclamation, "Message Error!"
    Exit Sub
End Sub

Private Sub cmd_refresh_Click()
If CommonDialog1.FileName <> "" Then
    Call fill_grid_excel_m(CommonDialog1.FileName)
End If
End Sub

Private Sub cmd_search_Click()
CommonDialog1.Filter = "XLS|*.xls"
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowOpen

If CommonDialog1.FileName <> "" Then
    Call fill_grid_excel_m(CommonDialog1.FileName)
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Frame1.Visible = False
    createGrid
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "Tanggal", 1200, lgAlignCenterCenter, lgDate, "yyyy-MM-dd", , , , , True
      .AddColumn "NIK", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Employee Name", 3500, , , , , , , , True
      .AddColumn "Jam Kerja", 1200, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Status", 500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Masuk", 1200, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
      .AddColumn "Keluar", 1200, lgAlignCenterCenter, lgDate, "hh:mm", , , , , , True
      .AddColumn "Terlambat", 1200, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Keterangan", 3000, , , , , , , , True
      .AddColumn "Shift", 1000, , , , , , , , False
      .AddColumn "Hari Libur", 1000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
End Sub

Private Sub CmdSave_Click()
Dim rsabsen As New ADODB.Recordset
Dim strsql As String
Dim clsAtt As New clsInsertAttendance
Dim i, j, dep, div, comp, Ttl, log, log_false, link, link_false As Integer

j = 0
dep = 0
div = 0
comp = 0
Ttl = 0
log = 0
log_false = 0
link = 0
link_false = 0

i = MsgBox("Are you sure to import all data?", vbOKCancel, headerMSG)
If Not i = vbOK Then Exit Sub
    
    Frame1.Visible = True
    
    If LynxGrid1.Rows > 0 Then
        Dim insertAtt As New clsInsertAtt_New
        Dim tgl_att As String, kdkaryawan As String, jam_masuk As String, jam_keluar As String, enroll_number As String
        
        DoEvents
        
        ProgressBar1.Value = 0
        
        With LynxGrid1
            For i = 0 To .Rows - 1
                
                ProgressBar1.Max = .Rows
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                Label5.Caption = "Saving Data, Please Wait......"
                Label1.Caption = "(" & .CellText(i, 1) & ") " & .CellText(i, 2)
                
                Select Case .CellText(i, 4)
                Case "SK", "PR", "S", "TH"
                    Dim abstatus As Integer
                    If .CellText(i, 4) = "SK" Then
                        abstatus = 2
                    ElseIf .CellText(i, 4) = "TH" Then
                        abstatus = 6
                    ElseIf .CellText(i, 4) = "PR" Then
                        abstatus = 0
                    ElseIf .CellText(i, 4) = "S" Then
                        abstatus = 1
                    End If
                End Select
            
                tgl_in = Format(.CellText(i, 0), "yyyy-MM-dd") & " " & Format(.CellText(i, 5), "hh:mm:ss")
                If Format(.CellText(i, 6), "hh:mm:ss") <= Format(.CellText(i, 5), "hh:mm:ss") Then
                    tgl_out = Format(DateAdd("d", 1, .CellText(i, 0)), "yyyy-MM-dd") & " " & Format(.CellText(i, 6), "hh:mm") & ":00"
                Else
                    tgl_out = Format(.CellText(i, 0), "yyyy-MM-dd") & " " & Format(.CellText(i, 6), "hh:mm") & ":00"
                End If
'                tgl_out = Format(.CellText(i, 0), "yyyy-MM-dd") & " " & Format(.CellText(i, 6), "hh:mm:ss")
                v_hari_libur = IIf(.CellText(i, 10) = "", 0, .CellText(i, 10))
                
                Dim v_employee_code As String
                strsql = "SELECT a.*, b.flag_ot FROM m_employee a join m_salary_standard b on a.employee_code = b.employee_code " _
                        & "WHERE a.nik = '" & .CellText(i, 1) & "' AND a.flag_active <> 0 order by b.number desc limit 1"
                rsabsen.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rsabsen.RecordCount > 0 Then
                    v_employee_code = rsabsen!employee_code
                    v_flag_ot = IIf(IsNull(rsabsen!flag_ot), 0, rsabsen!flag_ot)
                    v_nik = .CellText(i, 1)
                    v_employee_name = .CellText(i, 2)
                    v_nik_master = rsabsen!nik
                    v_employee_name_master = rsabsen!EMPLOYEE_NAME
                End If
                rsabsen.Close
                
                If v_nik <> v_nik_master Or UCase(v_employee_name) <> UCase(v_employee_name_master) Then
                    MsgBox "Ada Perbedaan Data " & Chr(13) & _
                        "* Data Excell : NIK : " & v_nik & " Dengan Nama : " & _
                        v_employee_name & Chr(13) & _
                        "* Data Master Karyawan : NIK : " & v_nik_master & " Dengan Nama : " & _
                        v_employee_name_master & Chr(13) & _
                        "Silahkan Update Data Excell Anda!", vbExclamation + vbOKOnly, "ERROR!!"
                    Frame1.Visible = False
                    Exit Sub
                End If
                
                Call totalJam
                Call totalLembur
                Call insertAtt.Insert_h_attendance(Format(.CellText(i, 0), "yyyy-MM-dd"), .CellText(i, 1), _
                    tgl_in, tgl_out, "0", .CellText(i, 1), .CellText(i, 8), .CellText(i, 9), .CellText(i, 4), _
                    v_tot_ot, v_15, v_2, v_3, v_4, v_tot_overtime, abstatus, v_employee_code, v_meal, v_trans, v_hari_libur, v_company_code)
            Next
        End With
    End If
    
    Frame1.Visible = False
    
'    MsgBox log & " log data are successfully import!" & vbCrLf _
'        & link & " link data are successfully import!" & vbCrLf _
'        & log_false & " log are unsuccessfully import!", vbInformation, headerMSG
    MsgBox "ALL Data are successfully import!", vbInformation
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    strsql = "Update temp_sal_proses set salary_proses = 0"
    CnG.Execute strsql
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
'    Call frm_mst_employee.load_data_company
'    Call frm_mst_employee.load_data_employee
'Unload Me
End Sub

Private Sub totalJam()
    Dim tglin As Date, tglout As Date
    Dim jambreak As Integer
    Dim tot_menit As Double
    Dim tgl_masuk As Date
    Dim tgl_keluar As Date
    Dim menit As Double
    Dim jam As Double
    Dim hour As Double
    
    Dim e As Long
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    
    strsql = "select company_code from m_employee WHERE nik = '" & LynxGrid1.CellText(i, 1) & "' " _
                    & "AND flag_active <> 0"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        v_company_code = rs!COMPANY_CODE
    End If
    rs.Close
    
    strsql = "Select start_time, end_time, min_break_in, max_break_out," & _
            "break_interval_minute FROM m_shift " & _
            "WHERE shift_code = '" & LynxGrid1.CellText(i, 9) & "' " & _
            "AND company_code = '" & v_company_code & "'"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
    Dim v_break_in As String
    Dim v_break_out As String
    
        v_break_in = Format(LynxGrid1.CellText(i, 0), "yyyy-MM-dd") & " " & Format(rs!min_break_in, "hh:nn:ss")
        v_break_out = Format(LynxGrid1.CellText(i, 0), "yyyy-MM-dd") & " " & Format(rs!max_break_out, "hh:nn:ss")
        v_start_time = Format(LynxGrid1.CellText(i, 0), "yyyy-MM-dd") & " " & Format(rs!start_time, "hh:nn:ss")
        v_end_time = Format(LynxGrid1.CellText(i, 0), "yyyy-MM-dd") & " " & Format(rs!end_time, "hh:nn:ss")
        v_interval = rs!break_interval_minute
    End If
    rs.Close
    
    jambreak = DateDiff("h", v_break_in, v_break_out) 'v_interval
    
    tgl_masuk = Format(tgl_in, "yyyy-MM-dd hh:nn:ss")
    tgl_keluar = Format(tgl_out, "yyyy-MM-dd hh:nn:ss")
    
    v_tot_work = DateDiff("h", v_start_time, v_end_time) - v_interval
    
    hour = (DateDiff("h", tgl_masuk, tgl_keluar)) - v_interval
    menit = (DateDiff("n", tgl_masuk, tgl_keluar)) / 60
    jam = roundDown(menit)
    
    e = (menit - jam) * 60
    If v_hari_libur = 0 Then
        If (hour - v_tot_work) < 1 Then
            If e <= 30 Then
                tot_menit = 0.5
            Else
                tot_menit = 1
            End If
            
            v_jmljam = jam - jambreak + tot_menit
        Else
            v_jmljam = jam - jambreak + (e / 60)
        End If
    Else
        If hour < 1 Then
            If e <= 30 Then
                tot_menit = 0.5
            Else
                tot_menit = 1
            End If
            
            v_jmljam = jam + tot_menit
        Else
            v_jmljam = jam + (e / 60)
        End If
    End If
    
    If v_jmljam < 0 Then
        v_jmljam = 0
    Else
        v_jmljam = v_jmljam
    End If

End Sub

Private Sub totalLembur()
'On Error Resume Next

v_tot_ot = Val(v_jmljam) - v_tot_work

If v_tot_ot = 0.5 Or v_tot_ot = 0 Then
    v_tot_ot = 0
Else
    If v_hari_libur = 0 Then
        v_tot_ot = Val(v_jmljam) - v_tot_work
    Else
        v_tot_ot = Val(v_jmljam)
    End If
End If
    
If v_hari_libur = 0 Then
    If v_flag_ot = 1 Then
        Select Case Val(v_tot_ot)
        Case Is = 1
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
        
        If Format(tgl_in, "hh:mm") <= "06:00" Or Format(tgl_out, "hh:mm") >= "22:00" Then
            v_trans = 1
        Else
            v_trans = 0
        End If
        
        If v_tot_ot >= 3 And v_tot_ot < 9 Then
            v_meal = 1
        ElseIf v_tot_ot >= 9 Then
            v_meal = 2
        Else
            v_meal = 0
        End If
    Else
        v_tot_ot = 0
        v_15 = 0
        v_2 = 0
        v_3 = 0
        v_4 = 0
        v_meal = 0
        v_trans = 0
    End If
Else
    If v_flag_ot = 1 Then
        v_tot_work = 0
        If v_tot_ot > 15 Then
            v_tot_ot = 15
        Else
            v_tot_ot = Val(v_jmljam)
        End If
            Select Case Val(v_tot_ot)
                Case Is <= 7
                    v_15 = 0
                    v_2 = Val(v_tot_ot)
                    v_3 = 0
                    v_4 = 0
                Case Is = 8
                    v_15 = 0
                    v_2 = 7
                    v_3 = 1
                    v_4 = 0
                Case Else
                    v_15 = 0
                    v_2 = 7
                    If (v_tot_ot - v_2) < 1 Then
                        v_3 = v_tot_ot - v_2
                    Else
                        v_3 = 1
                    End If
                    v_4 = Val(v_tot_ot) - Val(v_2) - Val(v_3)
            End Select
        
            v_trans = 2
            
            If v_tot_ot >= 3 And v_tot_ot < 11 Then
                v_meal = 1
            ElseIf v_tot_ot >= 11 Then
                v_meal = 2
            Else
                v_meal = 0
            End If
    Else
        v_tot_ot = 0
        v_15 = 0
        v_2 = 0
        v_3 = 0
        v_4 = 0
        v_meal = 0
        v_trans = 0
    End If
        
End If
    
'If v_flag_ot = 1 Then
'    Select Case Val(v_tot_ot)
'        Case Is = 1
'            v_15 = 1
'            v_2 = 0
'            v_3 = 0
'            v_4 = 0
'        Case Is > 1
'            v_15 = 1
'            v_2 = Val(v_tot_ot) - Val(v_15)
'            v_3 = 0
'            v_4 = 0
'        Case Else
'            v_15 = 0
'            v_2 = 0
'            v_3 = 0
'            v_4 = 0
'    End Select
'Else
'    v_15 = 0
'    v_2 = 0
'    v_3 = 0
'    v_4 = 0
'End If
'
'If v_flag_ot = 1 Then
'    If Format(LynxGrid1.CellText(i, 5), "hh:mm") < "06:00" Or Format(LynxGrid1.CellText(i, 6), "hh:mm") > "22:00" Then
'        v_trans = 1
'    Else
'        v_trans = 0
'    End If
'
'    If Val(v_jmljam) >= 3 And Val(v_jmljam) < 6 Then
'        v_meal = 1
'    ElseIf Val(v_jmljam) > 6 Then
'        v_meal = 2
'    Else
'        v_meal = 0
'    End If
'Else
'    v_meal = 0
'    v_trans = 0
'End If
    
v_tot_overtime = (1.5 * v_15) + (2 * v_2) + (3 * v_3) + (4 * v_4)
End Sub
