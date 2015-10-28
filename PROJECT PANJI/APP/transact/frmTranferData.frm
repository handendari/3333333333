VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_trans_tranfer_data 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EXPORT - IMPORT DATA"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11505
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   9090
      TabIndex        =   35
      Top             =   7410
      Width           =   1935
   End
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   5040
      Top             =   7530
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Pilih"
      Height          =   225
      Left            =   5700
      TabIndex        =   33
      Top             =   570
      Width           =   705
   End
   Begin VB.TextBox txt_department_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2970
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   29
      Top             =   510
      Width           =   2655
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   2970
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   24
      Top             =   120
      Width           =   2655
   End
   Begin prj_panji.vbButton btnBrowse 
      Height          =   315
      Left            =   4470
      TabIndex        =   23
      Top             =   6600
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
      MICON           =   "frmTranferData.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frmLoading 
      Height          =   555
      Left            =   3930
      TabIndex        =   19
      Top             =   4050
      Width           =   3735
      Begin VB.Label Label1 
         Caption         =   "Please wait...Loading Data Employee !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   210
         TabIndex        =   20
         Top             =   180
         Width           =   3465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   1260
      TabIndex        =   14
      Top             =   1470
      Width           =   4365
      Begin VB.OptionButton OptData 
         Caption         =   "Data Salary"
         Height          =   315
         Index           =   3
         Left            =   2100
         TabIndex        =   28
         Top             =   150
         Width           =   1665
      End
      Begin VB.OptionButton OptData 
         Caption         =   "Data Overtime"
         Height          =   315
         Index           =   2
         Left            =   2100
         TabIndex        =   27
         Top             =   510
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.OptionButton OptData 
         Caption         =   "Data Attendance"
         Height          =   315
         Index           =   1
         Left            =   2100
         TabIndex        =   18
         Top             =   330
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.OptionButton OptData 
         Caption         =   "Data Absensi "
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Width           =   1665
      End
      Begin VB.Label Label8 
         Caption         =   "dan Karyawan"
         Height          =   225
         Left            =   390
         TabIndex        =   34
         Top             =   390
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   1260
      TabIndex        =   13
      Top             =   930
      Width           =   4365
      Begin VB.OptionButton OptType 
         Caption         =   "Import"
         Height          =   285
         Index           =   1
         Left            =   2100
         TabIndex        =   16
         Top             =   180
         Width           =   2175
      End
      Begin VB.OptionButton OptType 
         Caption         =   "Export"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   180
         Width           =   2175
      End
   End
   Begin VB.CheckBox ChkAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10080
      TabIndex        =   4
      Top             =   2430
      Width           =   225
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      TabIndex        =   2
      Top             =   6600
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4050
      Top             =   7530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save File To"
      Filter          =   "*.ssd"
      InitDir         =   "C:\"
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   420
      TabIndex        =   1
      Top             =   6990
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   465
      Left            =   7050
      TabIndex        =   0
      Top             =   7410
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker_periode_from 
      Height          =   300
      Left            =   1260
      TabIndex        =   8
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   152633347
      CurrentDate     =   39278
      MaxDate         =   402133
      MinDate         =   36526
   End
   Begin MSComCtl2.DTPicker DTPicker_periode_to 
      Height          =   300
      Left            =   3930
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   152633347
      CurrentDate     =   39278
      MaxDate         =   402133
      MinDate         =   36526
   End
   Begin prj_panji.LynxGrid LynxGrid1 
      Height          =   3435
      Left            =   420
      TabIndex        =   22
      Top             =   2700
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   6059
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
      FocusRectMode   =   2
      FocusRectColor  =   9895934
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      AllowColumnResizing=   -1  'True
      Editable        =   -1  'True
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1260
      OleObjectBlob   =   "frmTranferData.frx":001C
      TabIndex        =   25
      Top             =   120
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1530
      Top             =   120
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_department 
      Height          =   375
      Left            =   1260
      OleObjectBlob   =   "frmTranferData.frx":1FDA
      TabIndex        =   30
      Top             =   510
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_department 
      Height          =   375
      Left            =   1530
      Top             =   540
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/ AREA"
      Height          =   255
      Left            =   60
      TabIndex        =   32
      Top             =   690
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT"
      Height          =   195
      Left            =   60
      TabIndex        =   31
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   390
      TabIndex        =   26
      Top             =   180
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Please wait..!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   420
      TabIndex        =   21
      Top             =   7320
      Width           =   3465
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA"
      Height          =   285
      Left            =   720
      TabIndex        =   12
      Top             =   1830
      Width           =   765
   End
   Begin VB.Label lblRange 
      BackStyle       =   0  'Transparent
      Caption         =   "s / d"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3150
      TabIndex        =   11
      Top             =   2310
      Width           =   615
   End
   Begin VB.Label lblPeriode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE"
      Height          =   195
      Left            =   450
      TabIndex        =   10
      Top             =   2340
      Width           =   720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE"
      Height          =   345
      Left            =   750
      TabIndex        =   7
      Top             =   1110
      Width           =   525
   End
   Begin VB.Label lblTitleGrid 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5670
      TabIndex        =   6
      Top             =   450
      Width           =   5295
   End
   Begin VB.Label lblAll 
      BackStyle       =   0  'Transparent
      Caption         =   "Select All"
      Height          =   255
      Left            =   10350
      TabIndex        =   5
      Top             =   2430
      Width           =   1035
   End
   Begin VB.Label lblFileTo 
      BackStyle       =   0  'Transparent
      Caption         =   "Export File To :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   420
      TabIndex        =   3
      Top             =   6180
      Width           =   2505
   End
End
Attribute VB_Name = "frm_trans_tranfer_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim conLite As New cConnection
Dim rsLite As cRecordset
Dim rs, rsemp, rsAtt As ADODB.Recordset
Dim strsql As String
Dim v_employee, v_attendance, v_bank, v_company As String
Dim v_department, v_division, v_level, v_location, v_title As String
Dim v_absent, v_duty, v_leave, v_general_leave, v_leave_periode As String
Dim v_salary_standard, v_other_income, v_other_expense As String
Dim v_loan, v_loan_detail, v_shift As String
Dim v_company_code As String

Private Sub cbo_periode_to_Change()
'Call periode_date_event
End Sub

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub Check2_Click()
If Check2.Value = 0 Then
    TDBCombo_department.Enabled = False
    txt_department_name.Enabled = False
    
    TDBCombo_department.Text = ""
    txt_department_name = ""
    
    If optType(0) And OptData(0) Then
        LynxGrid1.ClearAll
        Call createGridEmployee
        Call fillTableEmployee
    End If
Else
    TDBCombo_department.Enabled = True
    txt_department_name.Enabled = True
End If
End Sub

Private Sub ChkAll_Click()
Dim i As Integer
    Me.MousePointer = vbHourglass
    If LynxGrid1.Rows > 0 Then
        For i = 0 To LynxGrid1.Rows - 1
            If ChkAll.Value = 1 Then
                LynxGrid1.CellValue(i, 6) = 1
            Else
                LynxGrid1.CellValue(i, 6) = 0
            End If
        Next
    End If
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdRun_Click()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

If DTPicker_periode_from.Value = DTPicker_periode_to.Value And optType(0) And OptData(0) Then
    MsgBox "Pastikan Range Tanggal Terisi Dengan Benar!", vbExclamation, "Warning!"
    Exit Sub
End If

If Len(txtFileName.Text) = 0 Then
    MsgBox "Nama File Tidak Valid...", vbCritical
    Exit Sub
End If

If optType(0) Then ' Export Transfer Type

    If Right(txtFileName.Text, 4) = ".ssd" Then
        nmFile = txtFileName.Text
    Else
        nmFile = txtFileName.Text & ".ssd"
    End If

    If checkFile(nmFile) Then
        i = MsgBox("File Sudah Ada " & _
            "Tulis Ulang File ?", vbOKCancel, headerMSG)
        If i <> vbOK Then
            Exit Sub
        End If
        conLite.OpenDB nmFile
    Else
        conLite.CreateNewDB nmFile
    End If
    
    If OptData(0) Then ' Master Employee Data
        For aa = 0 To LynxGrid1.Rows - 1
            If LynxGrid1.CellValue(aa, 6) = True Then
                bb = bb + 1
            End If
        Next
         If bb <= 0 Then
            MsgBox "Data Belum Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        cmdRun.Enabled = False
        Call exportDataEmployee
        Call exportDataBank
        Call exportDataCompany
        Call exportDataDepartment
        Call exportDataDivision
        Call exportDataLevel
        Call exportDataLocation
        Call exportDataTitle
        Call exportDataAttendance
'        Call exportDataAbsent
'        Call exportDataDuty
'        Call exportDataLeave
        Call exportDataGeneralLeave
        Call exportDataLeavePeriode
        Call exportDataSalaryStandard
        Call exportDataOtherIncome
        Call exportDataOtherExpense
        Call exportDataLoan
        Call exportDataLoanDetail
        Call exportDataUser
        Call exportDataUserDetail
        Call exportDataUserAccess
        Call exportDataShift
        
        MsgBox "Data Berhasil Di Eksport dengan rincian : " & Chr(13) & _
            v_employee & Chr(13) & v_bank & Chr(13) & _
            v_company & Chr(13) & v_department & Chr(13) & v_division & Chr(13) & _
            v_level & Chr(13) & v_location & Chr(13) & v_title & Chr(13) & v_shift & Chr(13) & Chr(13) & _
            v_attendance & Chr(13) & Chr(13) & v_general_leave & Chr(13) & _
            v_leave_periode & Chr(13) & v_salary_standard & Chr(13) & _
            v_other_income & Chr(13) & v_other_expense & Chr(13) & _
            v_loan & Chr(13) & v_loan_detail, vbInformation, headerMSG
        cmdRun.Enabled = True
'    ElseIf OptData(1) Then ' Attendance Data
'        cmdRun.Enabled = False
'        Call exportDataAttendance
'        cmdRun.Enabled = True
'    ElseIf OptData(2) Then ' Overtime Data
'        cmdRun.Enabled = False
''        Call exportDataOvertime
'        cmdRun.Enabled = True
    ElseIf OptData(3) Then ' Salary Data
        cmdRun.Enabled = False
        Call exportDataHDSalary
        Call exportDataSalary
        cmdRun.Enabled = True
    End If
Else ' Import Transfer Type
    
    If Not Right(txtFileName.Text, 4) = ".ssd" Then
        MsgBox "Nama File Tidak Valid! Format Nama File Harus <namafile>.ssd..."
        Exit Sub
    End If
    
    If Not checkFile(txtFileName.Text) Then
        MsgBox "File Tidak Ditemukan..."
        Exit Sub
    End If
    
    cmdRun.Enabled = False
    
    If OptData(0) Then ' Master Employee Data
        Call importDataEmployee
        If v_company_code <> TDBCombo_company.Text Then
            MsgBox "File Untuk Perusahaan Ini Tidak Valid! Cek Ulang File...", vbExclamation, headerMSG
            cmdRun.Enabled = True
            Exit Sub
        End If
        Call importDataEmployee
        Call importDataBank
        Call importDataCompany
        Call importDataDepartment
        Call importDataDivision
        Call importDataLevel
        Call importDataLocation
        Call importDataTitle
        Call importDataAttendance
'        Call importDataAbsent
'        Call importDataDuty
'        Call importDataLeave
        Call importDataGeneralLeave
        Call importDataLeavePeriode
        Call importDataSalaryStandard
        Call importDataOtherIncome
        Call importDataOtherExpense
        Call importDataLoan
        Call importDataLoanDetail
        Call importDataUser
        Call importDataUserDetail
        Call importDataUserAccess
        Call importDataShift
        
        MsgBox "Data Berhasil Di Import dengan rincian : " & Chr(13) & _
            v_employee & Chr(13) & v_bank & Chr(13) & _
            v_company & Chr(13) & v_department & Chr(13) & v_division & Chr(13) & _
            v_level & Chr(13) & v_location & Chr(13) & v_title & Chr(13) & v_shift & Chr(13) & Chr(13) & _
            v_attendance & Chr(13) & Chr(13) & v_general_leave & Chr(13) & _
            v_leave_periode & Chr(13) & v_salary_standard & Chr(13) & _
            v_other_income & Chr(13) & v_other_expense & Chr(13) & _
            v_loan & Chr(13) & v_loan_detail, vbInformation, headerMSG
            
'    ElseIf OptData(1) Then ' Attendance Data
'        Call importDataAttendance
'    ElseIf OptData(2) Then
''        Call importDataOvertime
    ElseIf OptData(3) Then
        Call importDataHDSalary
        Call importDataSalary
    End If
    
    cmdRun.Enabled = True
End If

Set rss = Nothing

strsql = "UPDATE temp_sal_proses SET salary_proses = 1"
CnG.Execute strsql

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Adodc_company.ConnectionString = strConn
Adodc_department.ConnectionString = strConn

Call load_data_company

frmLoading.Visible = False
Label2.Visible = False

'Call createGridEmployee
'Call fillTableEmployee

optType(1) = True
OptData(0) = True
timer1.Enabled = True

Frame1.Enabled = False
Frame2.Enabled = False

End Sub

Private Sub OptType_Click(index As Integer)
   
    Call Option_Click
    
End Sub

Private Sub OptData_Click(index As Integer)
    
    Call Option_Click
    
End Sub

Private Sub Option_Click()
    
    If OptData(0) Then 'Employee List
        lblTitleGrid.Caption = " MASTER EMPLOYEE" & Chr(13) & "DAN ATTENDANCE LIST"
        Call mode_periode(False)
        If optType(0) Then
            Call mode_periode(True)
            lblFileTo.Caption = " Export File To : "
            LynxGrid1.Enabled = True
            Call createGridEmployee
                Frame1.Enabled = False
                Frame2.Enabled = False
                ChkAll.Visible = True
                lblAll.Visible = True
                ChkAll.Value = 0
            If fillTableEmployee() Then
                Frame1.Enabled = True
                Frame2.Enabled = True
            End If
            
'            If TDBCombo_company.Text = "GPN" Then
'                DTPicker_periode_from = Format(Now, "yyyy-MM-") & "01"
'                DTPicker_periode_to = Format(Now, "yyyy-MM-") & getEndDay(month(Now), year(Now))
'            Else
'                DTPicker_periode_from = Format(DateAdd("M", -1, Now), "yyyy-MM-") & "26"
'                DTPicker_periode_to = Format(Now, "yyyy-MM-") & "25"
'            End If

        Else
            Call mode_periode(False)
            lblFileTo.Caption = " Import File From : "
            If LynxGrid1.Rows > 0 Then
                LynxGrid1.ClearAll
            End If
            LynxGrid1.Enabled = False
            lblAll.Visible = False
            ChkAll.Visible = False
        End If
'    ElseIf OptData(1) Then 'Attendance List
'        lblTitleGrid.Caption = " ATTENDANCE DATA "
'        If LynxGrid1.Rows > 0 Then
'            LynxGrid1.ClearAll
'        End If
'        If OptType(0) Then
'            Call mode_periode(True)
'            lblFileTo.Caption = " Export File To : "
'            LynxGrid1.Enabled = True
'        Else
'            Call mode_periode(False)
'            lblFileTo.Caption = " Import File From : "
'            If LynxGrid1.Rows > 0 Then
'                LynxGrid1.ClearAll
'            End If
'            LynxGrid1.Enabled = False
'        End If
'        lblAll.Visible = False
'        ChkAll.Visible = False
'    ElseIf OptData(2) Then 'Overtime List
'        lblTitleGrid.Caption = " OVERTIME DATA "
'        If LynxGrid1.Rows > 0 Then
'            LynxGrid1.ClearAll
'        End If
'        If OptType(0) Then
'            Call mode_periode(True)
'            lblFileTo.Caption = " Export File To : "
'            LynxGrid1.Enabled = True
'        Else
'            Call mode_periode(False)
'            lblFileTo.Caption = " Import File From : "
'            If LynxGrid1.Rows > 0 Then
'                LynxGrid1.ClearAll
'            End If
'            LynxGrid1.Enabled = False
'        End If
'        lblAll.Visible = False
'        ChkAll.Visible = False
    ElseIf OptData(3) Then 'Salary List
        lblTitleGrid.Caption = " SALARY DATA "
        If LynxGrid1.Rows > 0 Then
            LynxGrid1.ClearAll
        End If
        If optType(0) Then
            Call mode_periode_salary(True)
            lblFileTo.Caption = " Export File To : "
            LynxGrid1.Enabled = True
        Else
            Call mode_periode_salary(False)
            lblFileTo.Caption = " Import File From : "
            If LynxGrid1.Rows > 0 Then
                LynxGrid1.ClearAll
            End If
            LynxGrid1.Enabled = False
        End If
        lblAll.Visible = False
        ChkAll.Visible = False
    End If
    
End Sub

Private Sub btnBrowse_Click()
If optType(0) Then
    CommonDialog1.ShowSave
Else
    CommonDialog1.ShowOpen
End If
If Len(CommonDialog1.FileName) > 0 Then
    txtFileName.Text = CommonDialog1.FileName
End If
End Sub

Private Sub createGridEmployee()

   With LynxGrid1
      .AddColumn "NIK", 1800, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Employee Name", 2700, , , , , , , , , True
      .AddColumn "Dept. Code", , , , , , , , , False
      .AddColumn "Dept. Name", 2500, , , , , , , , , True
      .AddColumn "Title Code", , , , , , , , , False
      .AddColumn "Title", 2500, , , , , , , , , True
      .AddColumn "Check", 800, lgAlignCenterCenter, lgBoolean, , , , , , , True
      .AddColumn "Emp. Code", , , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Function fillTableEmployee() As Boolean
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    
    'LynxGrid1.ClearAll
    fillTableEmployee = False
    frmLoading.Visible = True
    
    strsql = "SELECT nik,employee_name,department_code,department_name,title_code,title_name,employee_code " & _
        "FROM m_employee WHERE company_code = '" & TDBCombo_company.Text & "' AND " & _
        "CASE WHEN '" & Check2.Value & "' = 1 then flag_active = 1 AND department_code = '" & TDBCombo_department.Text & "' " & _
        "ELSE flag_active = 1 END " & _
        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
    
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME & _
                vbTab & rs!DEPARTMENT_CODE & vbTab & rs!department_name & _
                vbTab & rs!title_code & vbTab & rs!title_name & _
                vbTab & 0 & vbTab & rs!employee_code
            rs.MoveNext
        Wend
        frmLoading.Visible = False
        fillTableEmployee = True
    Else
        frmLoading.Visible = False
        Exit Function
    End If
End Function

Private Function checkFile(nmFile As String) As Boolean
    If Dir$(nmFile) <> "" Then
        checkFile = True
    Else
        checkFile = False
    End If
End Function

Private Sub mode_periode(str As Boolean)
    DTPicker_periode_from = Now
    DTPicker_periode_to = Now
    lblPeriode.Visible = str
    DTPicker_periode_from.Visible = str
    DTPicker_periode_from.CustomFormat = "yyyy-MM-dd"
    DTPicker_periode_to.Visible = str
    lblRange.Visible = str
    lblAll.Visible = IIf(str = True, False, True)
    ChkAll.Visible = IIf(str = True, False, True)
End Sub

Private Sub mode_periode_salary(str As Boolean)
    DTPicker_periode_from = Now
    DTPicker_periode_to = Now
    lblPeriode.Visible = str
    DTPicker_periode_from.Visible = str
    DTPicker_periode_from.Value = Format(DTPicker_periode_from, "yyyy-MM")
    DTPicker_periode_from.CustomFormat = "yyyy-MM"
    DTPicker_periode_to.Visible = False
    lblRange.Visible = False
    lblAll.Visible = IIf(str = True, False, True)
    ChkAll.Visible = IIf(str = True, False, True)
End Sub

Private Function check_exist_new_employee(empCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_employee = False

str_sql = "select count(employee_code) as rec_count from m_employee where employee_code = '" _
& Replace$(Trim$(empCode), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_employee = True
    Exit Function
End If
End Function

Private Function check_exist_new_company(compCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_company = False

str_sql = "select count(company_code) as rec_count from m_company where company_code = '" _
& Replace$(Trim$(compCode), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_company = True
    Exit Function
End If
End Function

Private Function check_exist_new_department(deptCode As String, compCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_department = False

str_sql = "select count(department_code) as rec_count from m_department where company_code = '" _
& Replace$(Trim$(compCode), Chr$(39), Chr$(96)) & "' AND department_code = '" _
& Replace$(Trim$(deptCode), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_department = True
    Exit Function
End If
End Function

Private Function check_exist_new_division(divCode As String, deptCode As String, compCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_division = False

str_sql = "select count(division_code) as rec_count from m_division where company_code = '" _
& Replace$(Trim$(compCode), Chr$(39), Chr$(96)) & "' AND department_code = '" _
& Replace$(Trim$(deptCode), Chr$(39), Chr$(96)) & "' AND division_code = '" _
& Replace$(Trim$(divCode), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_division = True
    Exit Function
End If
End Function

Private Function check_exist_new_location(locCode As String, compCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_location = False

str_sql = "select count(location_code) as rec_count from m_location where company_code = '" _
& Replace$(Trim$(compCode), Chr$(39), Chr$(96)) & "' AND location_code = '" _
& Replace$(Trim$(locCode), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_location = True
    Exit Function
End If
End Function

Private Function check_exist_new_bank(bankCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_bank = False

str_sql = "select count(bank_code) as rec_count from m_bank where bank_code = '" _
& Replace$(Trim$(bankCode), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_bank = True
    Exit Function
End If
End Function

Private Function check_exist_new_title(titleCode As String, compCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_title = False

str_sql = "select count(title_code) as rec_count from m_title where title_code = '" _
& Replace$(Trim$(titleCode), Chr$(39), Chr$(96)) & "' AND company_code = '" _
& Replace$(Trim$(compCode), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_title = True
    Exit Function
End If
End Function
Private Function check_exist_new_level(levelCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_level = False

str_sql = "select count(code) as rec_count from m_akses_level_group where code = '" _
& Replace$(Trim$(levelCode), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_level = True
    Exit Function
End If
End Function

Private Function check_exist_new_attendance(empCode As String, attDate As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_attendance = False

str_sql = "select count(employee_code and att_date) as rec_count from h_attendance where employee_code = '" & _
            Replace$(Trim$(empCode), Chr$(39), Chr$(96)) & "' AND att_date = '" & attDate & "' "
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_attendance = True
    Exit Function
End If
End Function

Private Function check_exist_new_absent(absNumber As Double) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_absent = False

str_sql = "select count(absent_number) as rec_count from t_absent where absent_number = '" & _
            Replace$(Trim$(absNumber), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_absent = True
    Exit Function
End If
End Function

Private Function check_exist_new_duty(dutyNumber As Double) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_duty = False

str_sql = "select count(duty_number) as rec_count from t_duty where duty_number = '" & _
            Replace$(Trim$(dutyNumber), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_duty = True
    Exit Function
End If
End Function

Private Function check_exist_new_leave(leaveNumber As Double) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_leave = False

str_sql = "select count(leave_number) as rec_count from t_leave where leave_number = '" & _
            Replace$(Trim$(leaveNumber), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_leave = True
    Exit Function
End If
End Function

Private Function check_exist_new_general_leave(general_leaveNumber As Double) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_general_leave = False

str_sql = "select count(general_leave_number) as rec_count from t_general_leave where general_leave_number = '" & _
            Replace$(Trim$(general_leaveNumber), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_general_leave = True
    Exit Function
End If
End Function

Private Function check_exist_new_leave_periode(empCode As String, start_periode As String, end_periode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_leave_periode = False

str_sql = "select count(employee_code) as rec_count from t_leave_periode where employee_code = '" & _
            Replace$(Trim$(empCode), Chr$(39), Chr$(96)) & "' AND date(start_periode) = '" & Format(start_periode, "yyyy-MM-dd") & "' AND " & _
            "date(end_periode) = '" & Format(end_periode, "yyyy-MM-dd") & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_leave_periode = True
    Exit Function
End If
End Function

Private Function check_exist_new_salary_standard(empCode As String, salary_date As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_salary_standard = False

str_sql = "select count(employee_code) as rec_count from m_salary_standard where employee_code = '" & _
            Replace$(Trim$(empCode), Chr$(39), Chr$(96)) & "' " & _
            "AND date(salary_date) = '" & Format(salary_date, "yyyy-MM-dd") & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_salary_standard = True
    Exit Function
End If
End Function

Private Function check_exist_new_other_income(noTrans As String, empCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_other_income = False

str_sql = "select count(notrans) as rec_count from t_employee_income where notrans = '" & _
            Replace$(Trim$(noTrans), Chr$(39), Chr$(96)) & "' " & _
            "AND employee_code = '" & empCode & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_other_income = True
    Exit Function
End If
End Function

Private Function check_exist_new_other_expense(noTrans As String, empCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_other_expense = False

str_sql = "select count(notrans) as rec_count from t_employee_expense where notrans = '" & _
            Replace$(Trim$(noTrans), Chr$(39), Chr$(96)) & "' " & _
            "AND employee_code = '" & empCode & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_other_expense = True
    Exit Function
End If
End Function

Private Function check_exist_new_tm_loan(empCode As String, tglLoan As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_tm_loan = False

str_sql = "select count(employee_code) as rec_count from tm_loan where employee_code = '" & _
            Replace$(Trim$(empCode), Chr$(39), Chr$(96)) & "' " & _
            "AND date(date) = '" & Format(tglLoan, "yyyy-MM-dd") & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_tm_loan = True
    Exit Function
End If
End Function

Private Function check_exist_new_td_loan(empCode As String, tglLoan As String, seq As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_td_loan = False

str_sql = "select count(employee_code) as rec_count from td_loan where employee_code = '" & _
            Replace$(Trim$(empCode), Chr$(39), Chr$(96)) & "' " & _
            "AND date(loan_date) = '" & Format(tglLoan, "yyyy-MM-dd") & "' " & _
            "AND sequence_number = '" & seq & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_td_loan = True
    Exit Function
End If
End Function

Private Function check_exist_new_m_user(userCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_m_user = False

str_sql = "select count(user_code) as rec_count from m_user where user_code = '" & _
            Replace$(Trim$(userCode), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_m_user = True
    Exit Function
End If
End Function

Private Function check_exist_new_t_user(lvlCode As String, sub_menu_code As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_t_user = False

str_sql = "select count(level_code) as rec_count from t_user where level_code = '" & _
            Replace$(Trim$(lvlCode), Chr$(39), Chr$(96)) & "' " & _
            "AND sub_menu_code = '" & sub_menu_code & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_t_user = True
    Exit Function
End If
End Function

Private Function check_exist_new_t_user_access(lvlCode As String, access_level_code As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_t_user_access = False

str_sql = "select count(level_code) as rec_count from t_user_access_level where level_code = '" & _
            Replace$(Trim$(lvlCode), Chr$(39), Chr$(96)) & "' " & _
            "AND access_level_code = '" & access_level_code & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_t_user_access = True
    Exit Function
End If
End Function

Private Function check_exist_new_shift(shiftCode As String, compCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_shift = False

str_sql = "select count(shift_code) as rec_count from m_shift where shift_code = '" & _
            Replace$(Trim$(shiftCode), Chr$(39), Chr$(96)) & "' " & _
            "AND company_code = '" & compCode & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_shift = True
    Exit Function
End If
End Function

Private Function check_exist_new_salary(month As String, empCode As String, salCode As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_new_salary = False

str_sql = "select count(month and employee_code and salary_code) as rec_count from h_salary where employee_code = '" & _
            Replace$(Trim$(empCode), Chr$(39), Chr$(96)) & "' AND month = '" & month & "' AND salary_code = '" & salCode & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_new_salary = True
    Exit Function
End If
End Function

Private Function check_exist_h_d_salary(month As String, COMPANY_CODE As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_h_d_salary = False

str_sql = "select count(company_code) as rec_count from h_d_salary where date(month) = '" & Format(month, "yyyy-MM-dd") & "' " & _
            "AND company_code = '" & COMPANY_CODE & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_exist_h_d_salary = True
    Exit Function
End If
End Function

Private Sub exportDataEmployee()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

    conLite.Execute "Create Table If Not Exists m_employee (" & _
                    "employee_code  TEXT(30) NOT NULL, employee_name  TEXT(50) DEFAULT NULL, employee_nick_name  TEXT(50) DEFAULT NULL, " & _
                    "division_code  TEXT(20) DEFAULT NULL, division_name  TEXT(50) DEFAULT NULL, department_code  TEXT(20) DEFAULT NULL, " & _
                    "department_name  TEXT(50) DEFAULT NULL, company_code  TEXT(20) DEFAULT NULL, company_name  TEXT(50) DEFAULT NULL, " & _
                    "date_of_birth  TEXT(30) DEFAULT NULL, place_of_birth  TEXT(100) DEFAULT NULL, sex  INTEGER DEFAULT NULL, " & _
                    "religion  INTEGER DEFAULT NULL, marital_status  INTEGER DEFAULT NULL, number_of_children  INTEGER DEFAULT NULL, " & _
                    "address  TEXT(100) DEFAULT NULL, email  TEXT(50) DEFAULT NULL, npwp  TEXT(50) DEFAULT NULL, " & _
                    "phone_number  TEXT(50) DEFAULT NULL, bank_account  TEXT(100) DEFAULT NULL, last_education_code  TEXT(20) DEFAULT NULL, " & _
                    "last_education_code_other  TEXT(20) DEFAULT NULL, last_education_name  TEXT(100) DEFAULT NULL, last_education_pass  TEXT(30) DEFAULT NULL,  " & _
                    "last_employment_name  TEXT(50) DEFAULT NULL, last_employment_date  TEXT(30) DEFAULT NULL, last_employment_title  TEXT(50) DEFAULT NULL," & _
                    "start_working  TEXT(30) DEFAULT NULL," & _
                    "date_of_appointment  TEXT(30) DEFAULT NULL, title_code  TEXT(20) DEFAULT NULL, title_name  TEXT(50) DEFAULT NULL, " & _
                    "level1  INTEGER DEFAULT NULL,level2  INTEGER DEFAULT NULL, " & _
                    "flag_shiftable  INTEGER DEFAULT NULL,flag_active  INTEGER DEFAULT NULL, " & _
                    "fathers_name  TEXT(50) DEFAULT NULL,mothers_name  TEXT(50) DEFAULT NULL, " & _
                    "child_number  INTEGER DEFAULT NULL,child_number_from  INTEGER DEFAULT NULL,description  TEXT(100) DEFAULT NULL, " & _
                    "end_working  TEXT(100) DEFAULT NULL,reason  TEXT(100) DEFAULT NULL,all_in  INTEGER DEFAULT NULL, " & _
                    "no_jamsostek  TEXT(30) DEFAULT NULL,level_code  INTEGER DEFAULT NULL," & _
                    "pertanggungan_pajak  INTEGER DEFAULT NULL,grade  TEXT(100) DEFAULT NULL,status_employee INTEGER DEFAULT NULL," & _
                    "prev_location  TEXT(50) DEFAULT NULL,curr_location  TEXT(50) DEFAULT NULL, " & _
                    "start_mc TEXT(50) DEFAULT NULL,end_mc TEXT(50) DEFAULT NULL, " & _
                    "nik  TEXT(20) NOT NULL,account_name TEXT(50) DEFAULT NULL,tgl_npwp TEXT(30) DEFAULT NULL,seq_no INTEGER DEFAULT NULL, " & _
                    "alamat_npwp TEXT(100) DEFAULT NULL, no_ktp TEXT(30) DEFAULT NULL, PRIMARY KEY(employee_code));"
    
    strsql = "INSERT INTO m_employee " & _
                "(employee_code,employee_name,employee_nick_name,division_code,division_name," & _
                "department_code,department_name,company_code,company_name,date_of_birth,place_of_birth," & _
                "sex,religion,marital_status,number_of_children,address,email,npwp,phone_number,bank_account," & _
                "last_education_code,last_education_code_other,last_education_name,last_education_pass,last_employment_name," & _
                "last_employment_date,last_employment_title,start_working,date_of_appointment,title_code,title_name," & _
                "level1,level2,flag_shiftable,flag_active,fathers_name,mothers_name," & _
                "child_number,child_number_from,description,end_working,reason," & _
                "all_in,no_jamsostek,level_code,pertanggungan_pajak,grade,status_employee," & _
                "prev_location,curr_location,start_mc,end_mc,nik,account_name,tgl_npwp,alamat_npwp,no_ktp,seq_no) " & _
                "Values " & _
                "(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    
    Set cmd = conLite.CreateCommand(strsql)
    
    bb = 0 'buat nyimpan jumlah data yang di select
    
    For aa = 0 To LynxGrid1.Rows - 1
        If LynxGrid1.CellValue(aa, 6) = True Then
            bb = bb + 1
        End If
    Next
    
    ProgressBar1.Max = bb
    ProgressBar1.Value = 0
    
    ProgressBar1.Visible = True
    Label2.Visible = True
    
    If LynxGrid1.Rows > 0 Then
        conLite.BeginTrans
        conLite.Execute ("DELETE FROM m_employee")
        For aa = 0 To LynxGrid1.Rows - 1
                        
            If LynxGrid1.CellValue(aa, 6) = True Then

                strsql = "SELECT * FROM m_employee WHERE employee_code = '" & LynxGrid1.CellText(aa, 7) & "' AND " & _
                        "CASE WHEN '" & Check2.Value & "' = 1 then company_code = '" & TDBCombo_company.Text & "' AND department_code = '" & TDBCombo_department.Text & "' " & _
                        "ELSE company_code = '" & TDBCombo_company.Text & "' END " & _
                        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
                Set rsemp = New ADODB.Recordset
                
                'If rsEmp.State = 1 Then rsEmp.Close
                rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                DoEvents
                ProgressBar1.Value = ProgressBar1.Value + 1
                Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                
                'Set rsLite = conLite.OpenRecordset(strSQl, True)
                cmd.SetText 1, rsemp!employee_code
                cmd.SetText 2, rsemp!EMPLOYEE_NAME
                cmd.SetText 3, IIf(IsNull(rsemp!employee_nick_name), "", rsemp!employee_nick_name)
                cmd.SetText 4, IIf(IsNull(rsemp!division_code), "", rsemp!division_code)
                cmd.SetText 5, IIf(IsNull(rsemp!division_name), "", rsemp!division_name)
                cmd.SetText 6, rsemp!DEPARTMENT_CODE
                cmd.SetText 7, rsemp!department_name
                cmd.SetText 8, rsemp!COMPANY_CODE
                cmd.SetText 9, IIf(IsNull(rsemp!company_name), "", rsemp!company_name)
                cmd.SetDate 10, IIf(IsNull(rsemp!date_of_birth) = True, 0, rsemp!date_of_birth)
                cmd.SetText 11, IIf(IsNull(rsemp!place_of_birth) = True, "", rsemp!place_of_birth)
                cmd.SetInt32 12, IIf(IsNull(rsemp!sex), 0, rsemp!sex)
                cmd.SetInt32 13, IIf(IsNull(rsemp!religion), 0, rsemp!religion)
                cmd.SetInt32 14, IIf(IsNull(rsemp!marital_status), 0, rsemp!marital_status)
                cmd.SetInt32 15, IIf(IsNull(rsemp!number_of_children), 0, rsemp!number_of_children)
                cmd.SetText 16, IIf(IsNull(rsemp!Address), "", rsemp!Address)
                cmd.SetText 17, IIf(IsNull(rsemp!email), "", rsemp!email)
                cmd.SetText 18, IIf(IsNull(rsemp!npwp), "", rsemp!npwp)
                cmd.SetText 19, IIf(IsNull(rsemp!phone_number), "", rsemp!phone_number)
                cmd.SetText 20, IIf(IsNull(rsemp!bank_account) = True, "", rsemp!bank_account)
                cmd.SetText 21, IIf(IsNull(rsemp!last_education_code), "", rsemp!last_education_code)
                cmd.SetText 22, IIf(IsNull(rsemp!last_education_code_other) = True, "", rsemp!last_education_code_other)
                cmd.SetText 23, IIf(IsNull(rsemp!last_education_name), "", rsemp!last_education_name)
                cmd.SetDate 24, IIf(IsNull(rsemp!last_education_pass) = True, 0, rsemp!last_education_pass)
                cmd.SetText 25, IIf(IsNull(rsemp!last_employment_name), "", rsemp!last_employment_name)
                cmd.SetDate 26, IIf(IsNull(rsemp!last_employment_date) = True, 0, rsemp!last_employment_date)
                cmd.SetText 27, IIf(IsNull(rsemp!last_employment_title), "", rsemp!last_employment_title)
                cmd.SetDate 28, IIf(IsNull(rsemp!start_working), 0, rsemp!start_working)
                cmd.SetDate 29, IIf(IsNull(rsemp!date_of_appointment) = True, 0, rsemp!date_of_appointment)
                cmd.SetText 30, IIf(IsNull(rsemp!title_code), "", rsemp!title_code)
                cmd.SetText 31, IIf(IsNull(rsemp!title_name), "", rsemp!title_name)
                cmd.SetInt32 32, IIf(IsNull(rsemp!level1), 0, rsemp!level1)
                cmd.SetInt32 33, IIf(IsNull(rsemp!level2), 0, rsemp!level2)
                cmd.SetInt32 34, IIf(IsNull(rsemp!flag_shiftable), 0, rsemp!flag_shiftable)
                cmd.SetInt32 35, IIf(IsNull(rsemp!flag_active), 0, rsemp!flag_active)
                cmd.SetText 36, IIf(IsNull(rsemp!fathers_name), "", rsemp!fathers_name)
                cmd.SetText 37, IIf(IsNull(rsemp!mothers_name), "", rsemp!mothers_name)
                cmd.SetInt32 38, IIf(IsNull(rsemp!child_number), 0, rsemp!child_number)
                cmd.SetInt32 39, IIf(IsNull(rsemp!child_number_from), 0, rsemp!child_number_from)
                cmd.SetText 40, IIf(IsNull(rsemp!Description), "", rsemp!Description)
                cmd.SetDate 41, IIf(IsNull(rsemp!end_working) = True, 0, rsemp!end_working)
                cmd.SetText 42, IIf(IsNull(rsemp!reason) = True, "", rsemp!reason)
                cmd.SetInt32 43, IIf(IsNull(rsemp!all_in), 0, rsemp!all_in)
                cmd.SetText 44, IIf(IsNull(rsemp!no_jamsostek) = True, "", rsemp!no_jamsostek)
                cmd.SetInt32 45, IIf(IsNull(rsemp!level_code), 0, rsemp!level_code)
                cmd.SetInt32 46, IIf(IsNull(rsemp!pertanggungan_pajak), 0, rsemp!pertanggungan_pajak)
                cmd.SetText 47, IIf(IsNull(rsemp!grade) = True, "", rsemp!grade)
                cmd.SetText 48, IIf(IsNull(rsemp!status_employee), 0, rsemp!status_employee)
                cmd.SetText 49, IIf(IsNull(rsemp!prev_location) = True, "", rsemp!prev_location)
                cmd.SetText 50, IIf(IsNull(rsemp!curr_location) = True, "", rsemp!curr_location)
                cmd.SetDate 51, IIf(IsNull(rsemp!start_mc) = True, 0, rsemp!start_mc)
                cmd.SetDate 52, IIf(IsNull(rsemp!end_mc) = True, 0, rsemp!end_mc)
                cmd.SetText 53, IIf(IsNull(rsemp!nik) = True, "", rsemp!nik)
                cmd.SetText 54, IIf(IsNull(rsemp!account_name) = True, "", rsemp!account_name)
                cmd.SetDate 55, IIf(IsNull(rsemp!tgl_npwp) = True, 0, rsemp!tgl_npwp)
                cmd.SetText 56, IIf(IsNull(rsemp!alamat_npwp) = True, "", rsemp!alamat_npwp)
                cmd.SetText 57, IIf(IsNull(rsemp!no_ktp) = True, "", rsemp!no_ktp)
                cmd.SetInt32 58, IIf(IsNull(rsemp!seq_no), 1, rsemp!seq_no)
                
                cmd.Execute
            End If
        Next
        conLite.CommitTrans
    End If
    rsemp.Close
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_employee = "Total Data Karyawan = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataCompany()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS m_company (" & _
                        "company_code TEXT(20) NOT NULL,company_name TEXT(100) NOT NULL," & _
                        "address TEXT(100) DEFAULT NULL,postal_code TEXT(30) DEFAULT NULL," & _
                        "city_name TEXT(50) NOT NULL,phone_number TEXT(20) NOT NULL," & _
                        "fax_number TEXT(20) DEFAULT NULL,web_address TEXT(20) DEFAULT NULL," & _
                        "email_address TEXT(50) NOT NULL,npwp TEXT(50) NOT NULL," & _
                        "pimpinan TEXT(50) DEFAULT NULL,pimpinan_npwp TEXT(50) DEFAULT NULL," & _
                        "npp TEXT(50) DEFAULT NULL, hari_kerja INTEGER DEFAULT NULL, PRIMARY KEY(company_code)); "
                        
        strsql = "INSERT INTO m_company " & _
                    "(company_code,company_name,address,postal_code," & _
                    "city_name,phone_number,fax_number,web_address," & _
                    "email_address,npwp,pimpinan,pimpinan_npwp,npp,hari_kerja) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT company_code,company_name,address,postal_code," & _
                    "city_name,phone_number,fax_number,web_address," & _
                    "email_address,npwp,pimpinan,pimpinan_npwp,npp,hari_kerja " & _
                 "FROM m_company WHERE company_code = '" & TDBCombo_company.Text & "'"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM m_company")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!COMPANY_CODE
                    cmd.SetText 2, rsAtt!company_name
                    cmd.SetText 3, rsAtt!Address
                    cmd.SetText 4, rsAtt!postal_code
                    cmd.SetText 5, rsAtt!city_name
                    cmd.SetText 6, rsAtt!phone_number
                    cmd.SetText 7, rsAtt!fax_number
                    cmd.SetText 8, rsAtt!web_address
                    cmd.SetText 9, rsAtt!email_address
                    cmd.SetText 10, rsAtt!npwp
                    cmd.SetText 11, rsAtt!pimpinan
                    cmd.SetText 12, rsAtt!pimpinan_npwp
                    cmd.SetText 13, IIf(IsNull(rsAtt!npp), "", rsAtt!npp)
                    cmd.SetInt32 14, IIf(IsNull(rsAtt!hari_kerja), 0, rsAtt!hari_kerja)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_company = "Total Data Perusahaan = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataDepartment()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS m_department (" & _
                        "company_code TEXT(20) NOT NULL,department_code TEXT(30) NOT NULL," & _
                        "department_name TEXT(100) DEFAULT NULL,description TEXT(100) DEFAULT NULL," & _
                        "PRIMARY KEY(department_code)); "
                        
        strsql = "INSERT INTO m_department " & _
                    "(company_code,department_code,department_name,description) " & _
                    "Values " & _
                    "(?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT company_code,department_code,department_name,description " & _
                 "FROM m_department WHERE " & _
                 "CASE WHEN '" & Check2.Value & "' = 1 then company_code = '" & TDBCombo_company.Text & "' AND department_code = '" & TDBCombo_department.Text & "' " & _
                 "ELSE company_code = '" & TDBCombo_company.Text & "' END"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM m_department")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!COMPANY_CODE
                    cmd.SetText 2, rsAtt!DEPARTMENT_CODE
                    cmd.SetText 3, rsAtt!department_name
                    cmd.SetText 4, IIf(IsNull(rsAtt!Description), "", rsAtt!Description)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_department = "Total Data Departemen / Area = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataDivision()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS m_division (" & _
                        "company_code TEXT(20) NOT NULL,department_code TEXT(30) NOT NULL," & _
                        "department_name TEXT(100) DEFAULT NULL,division_code TEXT(30) DEFAULT NULL," & _
                        "division_name TEXT(100) DEFAULT NULL,description TEXT(100) DEFAULT NULL," & _
                        "PRIMARY KEY(division_code)); "
                        
        strsql = "INSERT INTO m_division " & _
                    "(company_code,department_code,department_name,division_code,division_name,description) " & _
                    "Values " & _
                    "(?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT company_code,department_code,department_name,division_code,division_name,description " & _
                 "FROM m_division WHERE " & _
                 "CASE WHEN '" & Check2.Value & "' = 1 then company_code = '" & TDBCombo_company.Text & "' AND department_code = '" & TDBCombo_department.Text & "' " & _
                 "ELSE company_code = '" & TDBCombo_company.Text & "' END"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM m_division")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!COMPANY_CODE
                    cmd.SetText 2, rsAtt!DEPARTMENT_CODE
                    cmd.SetText 3, rsAtt!department_name
                    cmd.SetText 4, rsAtt!division_code
                    cmd.SetText 5, rsAtt!division_name
                    cmd.SetText 6, IIf(IsNull(rsAtt!Description), "", rsAtt!Description)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_division = "Total Data Divisi = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataLocation()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS m_location (" & _
                        "company_code TEXT(20) NOT NULL,location_code TEXT(30) NOT NULL," & _
                        "location_name TEXT(100) DEFAULT NULL,description TEXT(100) DEFAULT NULL," & _
                        "PRIMARY KEY(location_code)); "
                        
        strsql = "INSERT INTO m_location " & _
                    "(company_code,location_code,location_name,description) " & _
                    "Values " & _
                    "(?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT company_code,location_code,location_name,description " & _
                 "FROM m_location WHERE company_code = '" & TDBCombo_company.Text & "'"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM m_location")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!COMPANY_CODE
                    cmd.SetText 2, rsAtt!location_code
                    cmd.SetText 3, rsAtt!location_name
                    cmd.SetText 4, IIf(IsNull(rsAtt!Description), "", rsAtt!Description)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_location = "Total Data Lokasi = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataBank()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS m_bank (" & _
                        "bank_code TEXT(20) NOT NULL,bank_name TEXT(100) NOT NULL," & _
                        "description TEXT(100) DEFAULT NULL," & _
                        "PRIMARY KEY(bank_code)); "
                        
        strsql = "INSERT INTO m_bank " & _
                    "(bank_code,bank_name,description) " & _
                    "Values " & _
                    "(?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT bank_code,bank_name,description " & _
                 "FROM m_bank"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM m_bank")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!bank_code
                    cmd.SetText 2, rsAtt!bank_name
                    cmd.SetText 3, IIf(IsNull(rsAtt!Description), "", rsAtt!Description)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_bank = "Total Data Bank = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataTitle()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS m_title (" & _
                        "title_code TEXT(20) NOT NULL,title_name TEXT(100) NOT NULL," & _
                        "default_shiftable INTEGER DEFAULT NULL,level TEXT(20) NOT NULL," & _
                        "description TEXT(100) DEFAULT NULL,salary DECIMAL(15,2) NOT NULL," & _
                        "company_code TEXT(20) DEFAULT NULL,PRIMARY KEY(title_code)); "
                        
        strsql = "INSERT INTO m_title " & _
                    "(title_code,title_name,default_shiftable,level," & _
                    "description,salary,company_code) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT title_code,title_name,default_shiftable,level," & _
                    "description,salary,company_code " & _
                 "FROM m_title WHERE company_code = '" & TDBCombo_company.Text & "'"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM m_title")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!title_code
                    cmd.SetText 2, rsAtt!title_name
                    cmd.SetInt32 3, IIf(IsNull(rsAtt!default_shiftable) = True, 0, rsAtt!default_shiftable)
                    cmd.SetText 4, IIf(IsNull(rsAtt!Level), "", rsAtt!Level)
                    cmd.SetText 5, IIf(IsNull(rsAtt!Description), "", rsAtt!Description)
                    cmd.SetInt32 6, IIf(IsNull(rsAtt!salary) = True, 0, rsAtt!salary)
                    cmd.SetText 7, rsAtt!COMPANY_CODE
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_title = "Total Data Jabatan = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataLevel()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS m_akses_level_group (" & _
                        "code INTEGER NOT NULL,date_entry TEXT(30) DEFAULT NULL," & _
                        "date_edit TEXT(30) DEFAULT NULL,remark TEXT(100) NOT NULL," & _
                        "user_entry TEXT(20) DEFAULT NULL,user_edit TEXT(20) NOT NULL," & _
                        "name TEXT(50) DEFAULT NULL,PRIMARY KEY(code)); "
                        
        strsql = "INSERT INTO m_akses_level_group " & _
                    "(code,date_entry,date_edit,remark," & _
                    "user_entry,user_edit,name) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT code,date_entry,date_edit,remark," & _
                    "user_entry,user_edit,name " & _
                 "FROM m_akses_level_group"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM m_akses_level_group")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!Code
                    cmd.SetDate 2, IIf(IsNull(rsAtt!date_entry) = True, 0, rsAtt!date_entry)
                    cmd.SetDate 3, IIf(IsNull(rsAtt!date_edit) = True, 0, rsAtt!date_edit)
                    cmd.SetText 4, rsAtt!remark
                    cmd.SetText 5, rsAtt!user_entry
                    cmd.SetText 6, IIf(IsNull(rsAtt!user_edit), "", rsAtt!user_edit)
                    cmd.SetText 7, rsAtt!name
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_level = "Total Data Level = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataAttendance()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS h_attendance (" & _
                        "employee_code TEXT(20) NOT NULL, att_date TEXT(30) NOT NULL, ip_address TEXT(50) DEFAULT NULL," & _
                        "enrollnumber INTEGER DEFAULT NULL, shift_number TEXT(20) DEFAULT NULL,shift_code TEXT(20) DEFAULT NULL," & _
                        "start_time TEXT(30) DEFAULT NULL,end_time TEXT(30) DEFAULT NULL,time_in TEXT(30) DEFAULT NULL," & _
                        "time_out TEXT(30) DEFAULT NULL,time_in_diff TEXT(30) DEFAULT NULL,time_out_diff TEXT(30) DEFAULT NULL," & _
                        "flag_io INTEGER DEFAULT NULL,flag_present INTEGER DEFAULT NULL,flag_duty INTEGER DEFAULT NULL," & _
                        "flag_late INTEGER DEFAULT NULL,flag_early INTEGER DEFAULT NULL,absent_status TEXT(5) DEFAULT NULL," & _
                        "time_in_break TEXT(30) DEFAULT NULL,time_in_break_diff TEXT(30) DEFAULT NULL,flag_in_break_early INTEGER DEFAULT NULL," & _
                        "time_out_break TEXT(30) DEFAULT NULL,time_out_break_diff TEXT(30) DEFAULT NULL,flag_out_break_late INTEGER DEFAULT NULL," & _
                        "break_interval TEXT(30) DEFAULT NULL,break_ot datetime DEFAULT NULL,description TEXT(50) DEFAULT NULL," & _
                        "entry_date TEXT(30) DEFAULT NULL,edit_date TEXT(30) DEFAULT NULL,userinput TEXT(10) DEFAULT NULL," & _
                        "useredit TEXT(10) DEFAULT NULL,company_code TEXT(30) DEFAULT NULL," & _
                        "total_ot DECIMAL(15,2) DEFAULT NULL,satu_jam DECIMAL(15,2) DEFAULT NULL,dua_jam DECIMAL(15,2) DEFAULT NULL," & _
                        "tiga_jam DECIMAL(15,2) DEFAULT NULL,empat_jam DECIMAL(15,2) DEFAULT NULL," & _
                        "holiday INTEGER DEFAULT NULL,tot_overtime DECIMAL(15,2) DEFAULT NULL," & _
                        "flag_meal INTEGER DEFAULT NULL, flag_transport INTEGER DEFAULT NULL, PRIMARY KEY(employee_code, att_date)); "
                        
        strsql = "INSERT INTO h_attendance " & _
                    "(employee_code,att_date,ip_address,enrollnumber,shift_number," & _
                    "shift_code,start_time,end_time,time_in,time_out,time_in_diff," & _
                    "time_out_diff,flag_io,flag_present,flag_duty,flag_late,flag_early,absent_status," & _
                    "time_in_break,time_in_break_diff,flag_in_break_early,time_out_break,time_out_break_diff," & _
                    "flag_out_break_late,break_interval,break_ot,description,entry_date,edit_date,userinput," & _
                    "useredit,company_code,total_ot,satu_jam," & _
                    "dua_jam,tiga_jam,empat_jam,holiday,tot_overtime,flag_meal,flag_transport) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT a.employee_code,att_date,ip_address,enrollnumber,shift_number," & _
                    "shift_code,start_time,end_time,time_in,time_out,time_in_diff," & _
                    "time_out_diff,flag_io,flag_present,flag_duty,flag_late,flag_early,absent_status," & _
                    "time_in_break,time_in_break_diff,flag_in_break_early,time_out_break,time_out_break_diff," & _
                    "flag_out_break_late,break_interval,break_ot,a.description,entry_date,editdate,userinput," & _
                    "useredit,total_ot,15jam satu_jam," & _
                    "2jam dua_jam,3jam tiga_jam,4jam empat_jam,holiday,tot_overtime,flag_meal,flag_transport FROM h_attendance a JOIN m_employee b ON a.employee_code = b.employee_code WHERE DATE(att_date) >= '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " & _
                 "AND Date(att_date) <= '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "' AND " & _
                 "CASE WHEN '" & Check2.Value & "' = 1 then b.company_code = '" & TDBCombo_company.Text & "' AND b.department_code = '" & TDBCombo_department.Text & "' " & _
                 "ELSE b.company_code = '" & TDBCombo_company.Text & "' END"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = False
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM h_attendance")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                            
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!employee_code
                    cmd.SetDate 2, rsAtt!att_date
                    cmd.SetText 3, IIf(IsNull(rsAtt!ip_address) = True, "", rsAtt!ip_address)
                    cmd.SetText 4, IIf(IsNull(rsAtt!enrollnumber) = True, 0, rsAtt!enrollnumber)
                    cmd.SetText 5, IIf(IsNull(rsAtt!shift_number) = True, "", rsAtt!shift_number)
                    cmd.SetText 6, IIf(IsNull(rsAtt!shift_code) = True, 0, rsAtt!shift_code)
                    cmd.SetDate 7, IIf(IsNull(rsAtt!start_time) = True, 0, rsAtt!start_time)
                    cmd.SetDate 8, IIf(IsNull(rsAtt!end_time) = True, 0, rsAtt!end_time)
                    cmd.SetDate 9, IIf(IsNull(rsAtt!time_in) = True, 0, rsAtt!time_in)
                    cmd.SetDate 10, IIf(IsNull(rsAtt!time_out) = True, 0, rsAtt!time_out)
                    cmd.SetDate 11, IIf(IsNull(rsAtt!time_in_diff) = True, 0, rsAtt!time_in_diff)
                    cmd.SetDate 12, IIf(IsNull(rsAtt!time_out_diff) = True, 0, rsAtt!time_out_diff)
                    cmd.SetInt32 13, IIf(IsNull(rsAtt!flag_io) = True, 0, rsAtt!flag_io)
                    cmd.SetInt32 14, IIf(IsNull(rsAtt!flag_present) = True, 0, rsAtt!flag_present)
                    cmd.SetInt32 15, IIf(IsNull(rsAtt!flag_duty) = True, 0, rsAtt!flag_duty)
                    cmd.SetInt32 16, IIf(IsNull(rsAtt!flag_late) = True, 0, rsAtt!flag_late)
                    cmd.SetInt32 17, IIf(IsNull(rsAtt!flag_early) = True, 0, rsAtt!flag_early)
                    cmd.SetText 18, IIf(IsNull(rsAtt!absent_status) = True, "", rsAtt!absent_status)
                    cmd.SetDate 19, IIf(IsNull(rsAtt!time_in_break) = True, 0, rsAtt!time_in_break)
                    cmd.SetDate 20, IIf(IsNull(rsAtt!time_in_break_diff) = True, 0, rsAtt!time_in_break_diff)
                    cmd.SetInt32 21, IIf(IsNull(rsAtt!flag_in_break_early) = True, 0, rsAtt!flag_in_break_early)
                    cmd.SetDate 22, IIf(IsNull(rsAtt!time_out_break) = True, 0, rsAtt!time_out_break)
                    cmd.SetDate 23, IIf(IsNull(rsAtt!time_out_break_diff) = True, 0, rsAtt!time_out_break_diff)
                    cmd.SetInt32 24, IIf(IsNull(rsAtt!flag_out_break_late) = True, 0, rsAtt!flag_out_break_late)
                    cmd.SetDate 25, IIf(IsNull(rsAtt!break_interval) = True, 0, rsAtt!break_interval)
                    cmd.SetDate 26, IIf(IsNull(rsAtt!break_ot) = True, 0, rsAtt!break_ot)
                    cmd.SetText 27, IIf(IsNull(rsAtt!Description) = True, 0, rsAtt!Description)
                    cmd.SetDate 28, IIf(IsNull(rsAtt!entry_date) = True, 0, rsAtt!entry_date)
                    cmd.SetDate 29, IIf(IsNull(rsAtt!editdate) = True, 0, rsAtt!editdate)
                    cmd.SetText 30, IIf(IsNull(rsAtt!userInput) = True, 0, rsAtt!userInput)
                    cmd.SetText 31, IIf(IsNull(rsAtt!userEdit) = True, 0, rsAtt!userEdit)
                    cmd.SetText 32, Trim(TDBCombo_company.Text)
                    cmd.SetDouble 33, IIf(IsNull(rsAtt!total_OT) = True, 0, rsAtt!total_OT)
                    cmd.SetDouble 34, IIf(IsNull(rsAtt!satu_jam) = True, 0, rsAtt!satu_jam)
                    cmd.SetDouble 35, IIf(IsNull(rsAtt!dua_jam) = True, 0, rsAtt!dua_jam)
                    cmd.SetDouble 36, IIf(IsNull(rsAtt!tiga_jam) = True, 0, rsAtt!tiga_jam)
                    cmd.SetDouble 37, IIf(IsNull(rsAtt!empat_jam) = True, 0, rsAtt!empat_jam)
                    cmd.SetInt32 38, IIf(IsNull(rsAtt!holiday) = True, 0, rsAtt!holiday)
                    cmd.SetDouble 39, IIf(IsNull(rsAtt!tot_overtime) = True, 0, rsAtt!tot_overtime)
                    cmd.SetInt32 40, IIf(IsNull(rsAtt!flag_meal) = True, 0, rsAtt!flag_meal)
                    cmd.SetInt32 41, IIf(IsNull(rsAtt!flag_transport) = True, 0, rsAtt!flag_transport)
                    
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_attendance = "Total Data Kehadiran = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataAbsent()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS t_absent (" & _
                        "absent_number INTEGER NOT NULL,employee_code TEXT(20) DEFAULT NULL," & _
                        "absent_date_from TEXT(30) DEFAULT NULL,flag_date_to INTEGER NOT NULL," & _
                        "absent_date_to TEXT(30) DEFAULT NULL,absent_status INTEGER NOT NULL," & _
                        "description TEXT(100) DEFAULT NULL,entry_date TEXT(30) DEFAULT NULL," & _
                        "company_code TEXT(20) DEFAULT NULL,PRIMARY KEY(absent_number)); "
                        
        strsql = "INSERT INTO t_absent " & _
                    "(absent_number,employee_code,absent_date_from,flag_date_to," & _
                    "absent_date_to,absent_status,description,entry_date,company_code) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT absent_number,a.employee_code,absent_date_from,flag_date_to," & _
                    "absent_date_to,absent_status,a.description,entry_date " & _
                 "FROM t_absent a JOIN m_employee b ON a.employee_code = b.employee_code WHERE " & _
                 "(date(absent_date_from) between '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "') OR " & _
                 "(date(absent_date_to) between '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "') AND " & _
                 "CASE WHEN '" & Check2.Value & "' = 1 then b.company_code = '" & TDBCombo_company.Text & "' AND b.department_code = '" & TDBCombo_department.Text & "' " & _
                 "ELSE b.company_code = '" & TDBCombo_company.Text & "' END"
                         
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM t_absent")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetInt32 1, rsAtt!absent_number
                    cmd.SetText 2, rsAtt!employee_code
                    cmd.SetDate 3, IIf(IsNull(rsAtt!absent_date_from) = True, 0, rsAtt!absent_date_from)
                    cmd.SetInt32 4, rsAtt!flag_date_to
                    cmd.SetDate 5, IIf(IsNull(rsAtt!absent_date_to) = True, 0, rsAtt!absent_date_to)
                    cmd.SetInt32 6, rsAtt!absent_status
                    cmd.SetText 7, IIf(IsNull(rsAtt!Description), "", rsAtt!Description)
                    cmd.SetText 8, IIf(IsNull(rsAtt!entry_date) = True, 0, rsAtt!entry_date)
                    cmd.SetText 9, Trim(TDBCombo_company.Text)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_absent = "Total Data Absensi = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataDuty()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS t_duty (" & _
                        "duty_number INTEGER NOT NULL,employee_code TEXT(20) DEFAULT NULL," & _
                        "duty_date_from TEXT(30) DEFAULT NULL,flag_date_to INTEGER NOT NULL," & _
                        "duty_date_to TEXT(30) DEFAULT NULL," & _
                        "description TEXT(100) DEFAULT NULL,entry_date TEXT(30) DEFAULT NULL," & _
                        "company_code TEXT(20) DEFAULT NULL,PRIMARY KEY(duty_number)); "
                        
        strsql = "INSERT INTO t_duty " & _
                    "(duty_number,employee_code,duty_date_from,flag_date_to," & _
                    "duty_date_to,description,entry_date,company_code) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT duty_number,a.employee_code,duty_date_from,flag_date_to," & _
                    "duty_date_to,a.description,entry_date " & _
                 "FROM t_duty a JOIN m_employee b ON a.employee_code = b.employee_code WHERE " & _
                 "(date(duty_date_from) between '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "') OR " & _
                 "(date(duty_date_to) between '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "') AND " & _
                 "CASE WHEN '" & Check2.Value & "' = 1 then b.company_code = '" & TDBCombo_company.Text & "' AND b.department_code = '" & TDBCombo_department.Text & "' " & _
                 "ELSE b.company_code = '" & TDBCombo_company.Text & "' END"
                         
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM t_duty")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetInt32 1, rsAtt!duty_number
                    cmd.SetText 2, rsAtt!employee_code
                    cmd.SetDate 3, IIf(IsNull(rsAtt!duty_date_from) = True, 0, rsAtt!duty_date_from)
                    cmd.SetInt32 4, rsAtt!flag_date_to
                    cmd.SetDate 5, IIf(IsNull(rsAtt!duty_date_to) = True, 0, rsAtt!duty_date_to)
                    cmd.SetText 6, IIf(IsNull(rsAtt!Description), "", rsAtt!Description)
                    cmd.SetText 7, IIf(IsNull(rsAtt!entry_date) = True, 0, rsAtt!entry_date)
                    cmd.SetText 8, Trim(TDBCombo_company.Text)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_duty = "Total Data Tugas Dinas = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataLeave()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS t_leave (" & _
                        "leave_number INTEGER NOT NULL,employee_code TEXT(20) DEFAULT NULL," & _
                        "leave_type INTEGER DEFAULT NULL,flag_doctor_order INTEGER NOT NULL," & _
                        "leave_date_from TEXT(30) DEFAULT NULL,flag_date_to INTEGER NOT NULL," & _
                        "leave_date_to TEXT(30) DEFAULT NULL," & _
                        "description TEXT(100) DEFAULT NULL,entry_date TEXT(30) DEFAULT NULL," & _
                        "company_code TEXT(20) DEFAULT NULL,PRIMARY KEY(leave_number)); "
                        
        strsql = "INSERT INTO t_leave " & _
                    "(leave_number,employee_code,leave_type,flag_doctor_order,leave_date_from,flag_date_to," & _
                    "leave_date_to,description,entry_date,company_code) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT leave_number,a.employee_code,leave_type,flag_doctor_order,leave_date_from,flag_date_to," & _
                    "leave_date_to,a.description,entry_date " & _
                 "FROM t_leave a JOIN m_employee b ON a.employee_code = b.employee_code WHERE " & _
                 "(date(leave_date_from) between '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "') OR " & _
                 "(date(leave_date_to) between '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "') AND " & _
                 "CASE WHEN '" & Check2.Value & "' = 1 then b.company_code = '" & TDBCombo_company.Text & "' AND b.department_code = '" & TDBCombo_department.Text & "' " & _
                 "ELSE b.company_code = '" & TDBCombo_company.Text & "' END"
                         
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM t_leave")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetInt32 1, rsAtt!leave_number
                    cmd.SetText 2, rsAtt!employee_code
                    cmd.SetInt32 3, rsAtt!leave_type
                    cmd.SetInt32 4, rsAtt!flag_doctor_order
                    cmd.SetDate 5, IIf(IsNull(rsAtt!leave_date_from) = True, 0, rsAtt!leave_date_from)
                    cmd.SetInt32 6, rsAtt!flag_date_to
                    cmd.SetDate 7, IIf(IsNull(rsAtt!leave_date_to) = True, 0, rsAtt!leave_date_to)
                    cmd.SetText 8, IIf(IsNull(rsAtt!Description), "", rsAtt!Description)
                    cmd.SetText 9, IIf(IsNull(rsAtt!entry_date) = True, 0, rsAtt!entry_date)
                    cmd.SetText 10, Trim(TDBCombo_company.Text)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_leave = "Total Data Cuti = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataGeneralLeave()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS t_general_leave (" & _
                        "general_leave_number INTEGER NOT NULL,general_leave_date TEXT(30) DEFAULT NULL," & _
                        "description TEXT(100) DEFAULT NULL,entry_date TEXT(30) DEFAULT NULL," & _
                        "company_code TEXT(30) DEFAULT NULL,PRIMARY KEY(general_leave_number)); "
                        
        strsql = "INSERT INTO t_general_leave " & _
                    "(general_leave_number,general_leave_date,entry_date,company_code) " & _
                    "Values " & _
                    "(?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT general_leave_number,general_leave_date,entry_date,company_code " & _
                 "FROM t_general_leave WHERE company_code = '" & TDBCombo_company.Text & "' AND " & _
                 "date(general_leave_date) between '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "'"
                         
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM t_general_leave")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetInt32 1, rsAtt!general_leave_number
                    cmd.SetDate 2, IIf(IsNull(rsAtt!general_leave_date) = True, 0, rsAtt!general_leave_date)
                    cmd.SetText 3, IIf(IsNull(rsAtt!Description), "", rsAtt!Description)
                    cmd.SetText 4, IIf(IsNull(rsAtt!entry_date) = True, 0, rsAtt!entry_date)
                    cmd.SetText 5, IIf(IsNull(rsAtt!COMPANY_CODE) = True, "", rsAtt!COMPANY_CODE)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_general_leave = "Total Data Cuti Bersama = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataLeavePeriode()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS t_leave_periode (" & _
                        "employee_code TEXT(20) NOT NULL,start_working TEXT(30) DEFAULT NULL," & _
                        "start_periode TEXT(30) DEFAULT NULL,end_periode TEXT(30) NOT NULL," & _
                        "max_leave INTEGER DEFAULT NULL,actual_leave INTEGER NOT NULL," & _
                        "over_leave INTEGER DEFAULT NULL,flag_close INTEGER DEFAULT NULL," & _
                        "company_code TEXT(20) DEFAULT NULL,PRIMARY KEY(employee_code,start_periode,end_periode)); "
                        
        strsql = "INSERT INTO t_leave_periode " & _
                    "(employee_code,start_working,start_periode,end_periode,max_leave," & _
                    "actual_leave,over_leave,flag_close,company_code) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT a.employee_code,a.start_working,start_periode,end_periode,max_leave," & _
                    "actual_leave,over_leave,flag_close " & _
                 "FROM t_leave_periode a JOIN m_employee b ON a.employee_code = b.employee_code WHERE " & _
                 "(date(start_periode) between '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "') OR " & _
                 "(date(end_periode) between '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "') AND " & _
                 "CASE WHEN '" & Check2.Value & "' = 1 then b.company_code = '" & TDBCombo_company.Text & "' AND b.department_code = '" & TDBCombo_department.Text & "' " & _
                 "ELSE b.company_code = '" & TDBCombo_company.Text & "' END"
                         
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM t_leave_periode")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!employee_code
                    cmd.SetDate 2, IIf(IsNull(rsAtt!start_working) = True, 0, rsAtt!start_working)
                    cmd.SetDate 3, IIf(IsNull(rsAtt!start_periode) = True, 0, rsAtt!start_periode)
                    cmd.SetDate 4, IIf(IsNull(rsAtt!end_periode) = True, 0, rsAtt!end_periode)
                    cmd.SetInt32 5, rsAtt!max_leave
                    cmd.SetInt32 6, rsAtt!actual_leave
                    cmd.SetInt32 7, rsAtt!over_leave
                    cmd.SetInt32 8, rsAtt!flag_close
                    cmd.SetText 9, Trim(TDBCombo_company.Text)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    v_leave_periode = "Total Data Periode Cuti = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataSalaryStandard()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS m_salary_standard (" & _
                        "main_salary DECIMAL(15,2) DEFAULT NULL, staff_allowance DECIMAL(15,2) DEFAULT NULL," & _
                        "functional_allowance DECIMAL(15,2) DEFAULT NULL, phone_allowance DECIMAL(15,2) DEFAULT NULL," & _
                        "transport_allowance DECIMAL(15,2) DEFAULT NULL, other_allowance DECIMAL(15,2) DEFAULT NULL," & _
                        "presence_allowance DECIMAL(15,2) DEFAULT NULL,meal_allowance DECIMAL(15,2) DEFAULT NULL," & _
                        "spesial_allowance DECIMAL(15,2) DEFAULT NULL,employee_code TEXT(30) NOT NULL," & _
                        "driver_allowance DECIMAL(15,2) DEFAULT NULL,acting_allowance DECIMAL(15,2) DEFAULT NULL," & _
                        "skill_allowance DECIMAL(15,2) DEFAULT NULL,entry_date TEXT(30) DEFAULT NULL," & _
                        "user_entry TEXT(10) DEFAULT NULL, pph21_type TEXT(10) DEFAULT NULL,ptkp_type TEXT(10) DEFAULT NULL," & _
                        "jstk_type TEXT(10) DEFAULT NULL,flag_ot INTEGER DEFAULT NULL,salary_date TEXT(30) DEFAULT NULL," & _
                        "company_code TEXT(20) DEFAULT NULL," & _
                        "PRIMARY KEY(employee_code,salary_date)); "
                        
        strsql = "INSERT INTO m_salary_standard " & _
                    "(main_salary, staff_allowance,functional_allowance, phone_allowance," & _
                    "transport_allowance, other_allowance,presence_allowance," & _
                    "meal_allowance,spesial_allowance,employee_code,driver_allowance," & _
                    "acting_allowance,skill_allowance,entry_date,user_entry," & _
                    "pph21_type,ptkp_type,jstk_type,flag_ot,salary_date,company_code)" & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT a.* FROM m_salary_standard a join m_employee b on a.employee_code = b.employee_code " _
                & "WHERE (date(a.salary_date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "') AND " _
                & "CASE WHEN '" & Check2.Value & "' = 1 then b.company_code = '" & TDBCombo_company.Text & "' AND b.department_code = '" & TDBCombo_department.Text & "' " _
                & "ELSE b.company_code = '" & TDBCombo_company.Text & "' END"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM m_salary_standard")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetDouble 1, IIf(IsNull(rsAtt!main_salary) = True, 0, rsAtt!main_salary)
                    cmd.SetDouble 2, IIf(IsNull(rsAtt!staff_allowance) = True, 0, rsAtt!staff_allowance)
                    cmd.SetDouble 3, IIf(IsNull(rsAtt!functional_allowance) = True, 0, rsAtt!functional_allowance)
                    cmd.SetDouble 4, IIf(IsNull(rsAtt!phone_allowance) = True, 0, rsAtt!phone_allowance)
                    cmd.SetDouble 5, IIf(IsNull(rsAtt!transport_allowance) = True, 0, rsAtt!transport_allowance)
                    cmd.SetDouble 6, IIf(IsNull(rsAtt!other_allowance) = True, 0, rsAtt!other_allowance)
                    cmd.SetDouble 7, IIf(IsNull(rsAtt!presence_allowance) = True, 0, rsAtt!presence_allowance)
                    cmd.SetDouble 8, IIf(IsNull(rsAtt!meal_allowance) = True, 0, rsAtt!meal_allowance)
                    cmd.SetDouble 9, IIf(IsNull(rsAtt!special_allowance) = True, 0, rsAtt!special_allowance)
                    cmd.SetText 10, rsAtt!employee_code
                    cmd.SetDouble 11, IIf(IsNull(rsAtt!driver_allowance) = True, 0, rsAtt!driver_allowance)
                    cmd.SetDouble 12, IIf(IsNull(rsAtt!acting_allowance) = True, 0, rsAtt!acting_allowance)
                    cmd.SetDouble 13, IIf(IsNull(rsAtt!skill_allowance) = True, 0, rsAtt!skill_allowance)
                    cmd.SetDate 14, IIf(IsNull(rsAtt!entry_date) = True, 0, rsAtt!entry_date)
                    cmd.SetText 15, IIf(IsNull(rsAtt!user_entry) = True, "", rsAtt!user_entry)
                    cmd.SetText 16, IIf(IsNull(rsAtt!pph21_type) = True, "", rsAtt!pph21_type)
                    cmd.SetText 17, IIf(IsNull(rsAtt!ptkp_type) = True, "", rsAtt!ptkp_type)
                    cmd.SetText 18, IIf(IsNull(rsAtt!jstk_type) = True, "", rsAtt!jstk_type)
                    cmd.SetInt32 19, IIf(IsNull(rsAtt!flag_ot) = True, 0, rsAtt!flag_ot)
                    cmd.SetDate 20, IIf(IsNull(rsAtt!salary_date) = True, 0, rsAtt!salary_date)
                    cmd.SetText 21, Trim(TDBCombo_company.Text)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    v_salary_standard = "Total Data Salary Standard = " & ProgressBar1.Value & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataOtherIncome()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS t_employee_income (" & _
                        "nomer INTEGER DEFAULT NULL,notrans TEXT(50) NOT NULL,tgltrans TEXT(20) DEFAULT NULL," & _
                        "employee_code TEXT(50) NOT NULL,jmlPotong DECIMAL(15,2) DEFAULT NULL,remark TEXT(30) DEFAULT NULL," & _
                        "userinput TEXT(30) DEFAULT NULL,tglinput TEXT(30) DEFAULT NULL," & _
                        "useredit TEXT(30) DEFAULT NULL,tgledit TEXT(30) DEFAULT NULL,flag_other_income INTEGER DEFAULT NULL," & _
                        "company_code TEXT(30) DEFAULT NULL, " & _
                        "PRIMARY KEY(notrans,employee_code)); "
                        
        strsql = "INSERT INTO t_employee_income " & _
                    "(nomer,notrans,tgltrans,employee_code,jmlPotong,remark," & _
                    "userinput,tglinput,useredit,tgledit,flag_other_income,company_code) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT a.* FROM t_employee_income a join m_employee b on a.employee_code = b.employee_code " _
                & "WHERE (a.tgltrans BETWEEN '" & DTPicker_periode_from.Value & "' AND '" & DTPicker_periode_to.Value & "') AND " _
                & "CASE WHEN '" & Check2.Value & "' = 1 then b.company_code = '" & TDBCombo_company.Text & "' AND b.department_code = '" & TDBCombo_department.Text & "' " _
                & "ELSE b.company_code = '" & TDBCombo_company.Text & "' END"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM t_employee_income")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetInt32 1, rsAtt!nomer
                    cmd.SetText 2, rsAtt!noTrans
                    cmd.SetDate 3, IIf(IsNull(rsAtt!tgltrans) = True, 0, rsAtt!tgltrans)
                    cmd.SetText 4, rsAtt!employee_code
                    cmd.SetDouble 5, IIf(IsNull(rsAtt!jmlpotong) = True, 0, rsAtt!jmlpotong)
                    cmd.SetText 6, IIf(IsNull(rsAtt!remark) = True, "", rsAtt!remark)
                    cmd.SetText 7, IIf(IsNull(rsAtt!userInput) = True, "", rsAtt!userInput)
                    cmd.SetDate 8, IIf(IsNull(rsAtt!tglinput) = True, 0, rsAtt!tglinput)
                    cmd.SetText 9, IIf(IsNull(rsAtt!userEdit) = True, "", rsAtt!userEdit)
                    cmd.SetDate 10, IIf(IsNull(rsAtt!tglEdit) = True, 0, rsAtt!tglEdit)
                    cmd.SetInt32 11, IIf(IsNull(rsAtt!flag_other_income) = True, 0, rsAtt!flag_other_income)
                    cmd.SetText 12, Trim(TDBCombo_company.Text)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    
    v_other_income = "Total Data Other Income = " & ProgressBar1.Value & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataOtherExpense()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS t_employee_expense (" & _
                        "nomer INTEGER DEFAULT NULL,notrans TEXT(50) NOT NULL,tgltrans TEXT(20) DEFAULT NULL," & _
                        "employee_code TEXT(50) NOT NULL,jmlPotong DECIMAL(15,2) DEFAULT NULL,remark TEXT(30) DEFAULT NULL," & _
                        "userinput TEXT(30) DEFAULT NULL,tglinput TEXT(30) DEFAULT NULL," & _
                        "useredit TEXT(30) DEFAULT NULL,tgledit TEXT(30) DEFAULT NULL,flag_other_expense INTEGER DEFAULT NULL," & _
                        "company_code TEXT(30) DEFAULT NULL, " & _
                        "PRIMARY KEY(notrans,employee_code)); "
                        
        strsql = "INSERT INTO t_employee_expense " & _
                    "(nomer,notrans,tgltrans,employee_code,jmlPotong,remark," & _
                    "userinput,tglinput,useredit,tgledit,flag_other_expense,company_code) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT a.* FROM t_employee_expense a join m_employee b on a.employee_code = b.employee_code " _
                & "WHERE (a.tgltrans BETWEEN '" & DTPicker_periode_from.Value & "' AND '" & DTPicker_periode_to.Value & "') AND " _
                & "CASE WHEN '" & Check2.Value & "' = 1 then b.company_code = '" & TDBCombo_company.Text & "' AND b.department_code = '" & TDBCombo_department.Text & "' " _
                & "ELSE b.company_code = '" & TDBCombo_company.Text & "' END"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM t_employee_expense")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetInt32 1, rsAtt!nomer
                    cmd.SetText 2, rsAtt!noTrans
                    cmd.SetDate 3, IIf(IsNull(rsAtt!tgltrans) = True, 0, rsAtt!tgltrans)
                    cmd.SetText 4, rsAtt!employee_code
                    cmd.SetDouble 5, IIf(IsNull(rsAtt!jmlpotong) = True, 0, rsAtt!jmlpotong)
                    cmd.SetText 6, IIf(IsNull(rsAtt!remark) = True, "", rsAtt!remark)
                    cmd.SetText 7, IIf(IsNull(rsAtt!userInput) = True, "", rsAtt!userInput)
                    cmd.SetDate 8, IIf(IsNull(rsAtt!tglinput) = True, 0, rsAtt!tglinput)
                    cmd.SetText 9, IIf(IsNull(rsAtt!userEdit) = True, "", rsAtt!userEdit)
                    cmd.SetDate 10, IIf(IsNull(rsAtt!tglEdit) = True, 0, rsAtt!tglEdit)
                    cmd.SetInt32 11, IIf(IsNull(rsAtt!flag_other_expense) = True, 0, rsAtt!flag_other_expense)
                    cmd.SetText 12, Trim(TDBCombo_company.Text)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    
    v_other_expense = "Total Data Other Expense = " & ProgressBar1.Value & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataLoan()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS tm_loan (" & _
                        "date TEXT(30) NOT NULL, employee_code TEXT(30) NOT NULL,loan_value DECIMAL(15,2) DEFAULT NULL," & _
                        "loan_interest DECIMAL(15,2) DEFAULT NULL,loan_total DECIMAL(15,2) DEFAULT NULL,installment_time INTEGER DEFAULT NULL," & _
                        "installment_value DECIMAL(15,2) DEFAULT NULL,installment_start TEXT(30) DEFAULT NULL," & _
                        "installment_end DECIMAL(15,2) DEFAULT NULL,description TEXT(30) DEFAULT NULL," & _
                        "company_code TEXT(30) DEFAULT NULL,employee_name TEXT(50) DEFAULT NULL," & _
                        "PRIMARY KEY(employee_code,date)); "
                        
        strsql = "INSERT INTO tm_loan " & _
                    "(date,employee_code,loan_value,loan_interest," & _
                    "loan_total,installment_time,installment_value," & _
                    "installment_start,installment_end,description,company_code,employee_name) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT a.* FROM tm_loan a join m_employee b on a.employee_code = b.employee_code " _
                & "WHERE (a.date BETWEEN '" & DTPicker_periode_from.Value & "' AND '" & DTPicker_periode_to.Value & "') AND " _
                & "CASE WHEN '" & Check2.Value & "' = 1 then b.company_code = '" & TDBCombo_company.Text & "' AND b.department_code = '" & TDBCombo_department.Text & "' " _
                & "ELSE b.company_code = '" & TDBCombo_company.Text & "' END"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM tm_loan")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetDate 1, rsAtt!Date
                    cmd.SetText 2, rsAtt!employee_code
                    cmd.SetDouble 3, IIf(IsNull(rsAtt!loan_value) = True, 0, rsAtt!loan_value)
                    cmd.SetDouble 4, IIf(IsNull(rsAtt!loan_interest) = True, 0, rsAtt!loan_interest)
                    cmd.SetDouble 5, IIf(IsNull(rsAtt!loan_total) = True, 0, rsAtt!loan_total)
                    cmd.SetInt32 6, IIf(IsNull(rsAtt!installment_time) = True, 0, rsAtt!installment_time)
                    cmd.SetDouble 7, IIf(IsNull(rsAtt!installment_value) = True, 0, rsAtt!installment_value)
                    cmd.SetDate 8, IIf(IsNull(rsAtt!installment_start) = True, 0, rsAtt!installment_start)
                    cmd.SetDate 9, IIf(IsNull(rsAtt!installment_end) = True, 0, rsAtt!installment_end)
                    cmd.SetText 10, IIf(IsNull(rsAtt!Description) = True, "", rsAtt!Description)
                    cmd.SetText 11, Trim(TDBCombo_company.Text)
                    cmd.SetText 12, IIf(IsNull(rsAtt!EMPLOYEE_NAME) = True, "", rsAtt!EMPLOYEE_NAME)
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    
    v_loan = "Total Data Loan = " & ProgressBar1.Value & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataLoanDetail()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS td_loan (" & _
                        "date TEXT(30) NOT NULL,employee_code TEXT(30) NOT NULL,salary_item_code TEXT(30) DEFAULT NULL," & _
                        "sequence_number INTEGER NOT NULL,employee_name TEXT(50) DEFAULT NULL,salary_item_name TEXT(50) DEFAULT NULL," & _
                        "installment_month TEXT(30) DEFAULT NULL,installment_ammount DECIMAL(15,2) DEFAULT NULL," & _
                        "installment_pay DECIMAL(15,2) DEFAULT NULL,installment_pay_date TEXT(30) DEFAULT NULL," & _
                        "flag_paid INTEGER DEFAULT NULL,company_code TEXT(30) DEFAULT NULL,PRIMARY KEY(date,employee_code,sequence_number)); "
                        
        strsql = "INSERT INTO td_loan " & _
                        "(date,employee_code,salary_item_code,sequence_number,employee_name," & _
                        "salary_item_name,installment_month,installment_ammount," & _
                        "installment_pay,installment_pay_date,flag_paid,company_code) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT a.* FROM td_loan a join m_employee b on a.employee_code = b.employee_code " _
                & "WHERE (a.date BETWEEN '" & DTPicker_periode_from.Value & "' AND '" & DTPicker_periode_to.Value & "') AND " _
                & "CASE WHEN '" & Check2.Value & "' = 1 then b.company_code = '" & TDBCombo_company.Text & "' AND b.department_code = '" & TDBCombo_department.Text & "' " _
                & "ELSE b.company_code = '" & TDBCombo_company.Text & "' END"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM td_loan")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetDate 1, rsAtt!Date
                    cmd.SetText 2, rsAtt!employee_code
                    cmd.SetText 3, IIf(IsNull(rsAtt!salary_item_code) = True, "", rsAtt!loan_value)
                    cmd.SetInt32 4, rsAtt!sequence_number
                    cmd.SetText 5, IIf(IsNull(rsAtt!EMPLOYEE_NAME) = True, "", rsAtt!EMPLOYEE_NAME)
                    cmd.SetText 6, IIf(IsNull(rsAtt!salary_item_name) = True, "", rsAtt!salary_item_name)
                    cmd.SetDate 7, IIf(IsNull(rsAtt!installment_month) = True, 0, rsAtt!installment_month)
                    cmd.SetDouble 8, IIf(IsNull(rsAtt!installment_amount) = True, 0, rsAtt!installment_amount)
                    cmd.SetDouble 9, IIf(IsNull(rsAtt!installment_pay) = True, 0, rsAtt!installment_pay_date)
                    cmd.SetDate 10, IIf(IsNull(rsAtt!installment_pay_date) = True, 0, rsAtt!installment_pay_date)
                    cmd.SetInt32 11, IIf(IsNull(rsAtt!flag_paid) = True, 0, rsAtt!flag_paid)
                    cmd.SetText 12, Trim(TDBCombo_company.Text)
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    
    v_loan_detail = "Total Data Loan Detail = " & ProgressBar1.Value & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataUser()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS m_user (" & _
                        "user_code TEXT(30) NOT NULL,company_code TEXT(30) DEFAULT NULL," & _
                        "user_name TEXT(50) DEFAULT NULL,user_pass TEXT(100) DEFAULT NULL," & _
                        "user_pass_key TEXT(50) DEFAULT NULL,user_level INTEGER DEFAULT NULL," & _
                        "flag_company_access INTEGER DEFAULT NULL,employee_code TEXT(30) DEFAULT NULL," & _
                        "flag_user INTEGER DEFAULT NULL,PRIMARY KEY(user_code)); "
                        
        strsql = "INSERT INTO m_user " & _
                    "(user_code,company_code,user_name,user_pass," & _
                    "user_pass_key,user_level,flag_company_access,employee_code," & _
                    "flag_user) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT user_code,company_code,user_name,user_pass," & _
                    "user_pass_key,user_level,flag_company_access,employee_code," & _
                    "flag_user " & _
                 "FROM m_user"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM m_user")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!user_code
                    cmd.SetText 2, IIf(IsNull(rsAtt!COMPANY_CODE), "", rsAtt!COMPANY_CODE)
                    cmd.SetText 3, IIf(IsNull(rsAtt!USER_NAME), "", rsAtt!USER_NAME)
                    cmd.SetText 4, IIf(IsNull(rsAtt!USER_PASS), "", rsAtt!USER_PASS)
                    cmd.SetText 5, IIf(IsNull(rsAtt!user_pass_key), "", rsAtt!user_pass_key)
                    cmd.SetInt32 6, IIf(IsNull(rsAtt!user_level), 0, rsAtt!user_level)
                    cmd.SetInt32 7, IIf(IsNull(rsAtt!flag_company_access), 0, rsAtt!flag_company_access)
                    cmd.SetText 8, IIf(IsNull(rsAtt!employee_code), "", rsAtt!employee_code)
                    cmd.SetInt32 9, IIf(IsNull(rsAtt!flag_user), 0, rsAtt!flag_user)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
'    v_company = "Total Data Perusahaan = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataUserDetail()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS t_user (" & _
                        "level_code TEXT(30) NOT NULL,sub_menu_code TEXT(30) NOT NULL," & _
                        "sub_menu_name TEXT(50) DEFAULT NULL,menu_code TEXT(50) DEFAULT NULL," & _
                        "menu_name TEXT(50) DEFAULT NULL,form_name TEXT(50) DEFAULT NULL," & _
                        "form_title TEXT(50) DEFAULT NULL,allow_read INTEGER DEFAULT NULL," & _
                        "allow_add INTEGER DEFAULT NULL,allow_edit INTEGER DEFAULT NULL," & _
                        "allow_delete INTEGER DEFAULT NULL,allow_post INTEGER DEFAULT NULL," & _
                        "allow_print INTEGER DEFAULT NULL,PRIMARY KEY(level_code,sub_menu_code)); "
                        
        strsql = "INSERT INTO t_user " & _
                    "(level_code,sub_menu_code,sub_menu_name,menu_code," & _
                    "menu_name,form_name,form_title,allow_read," & _
                    "allow_add,allow_edit,allow_delete,allow_post," & _
                    "allow_print) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT level_code,sub_menu_code,sub_menu_name,menu_code," & _
                    "menu_name,form_name,form_title,allow_read," & _
                    "allow_add,allow_edit,allow_delete,allow_post," & _
                    "allow_print " & _
                 "FROM t_user"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM t_user")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!level_code
                    cmd.SetText 2, rsAtt!sub_menu_code
                    cmd.SetText 3, IIf(IsNull(rsAtt!sub_menu_name), "", rsAtt!sub_menu_name)
                    cmd.SetText 4, IIf(IsNull(rsAtt!menu_code), "", rsAtt!menu_code)
                    cmd.SetText 5, IIf(IsNull(rsAtt!menu_name), "", rsAtt!menu_name)
                    cmd.SetText 6, IIf(IsNull(rsAtt!form_name), "", rsAtt!form_name)
                    cmd.SetText 7, IIf(IsNull(rsAtt!form_title), "", rsAtt!form_title)
                    cmd.SetInt32 8, IIf(IsNull(rsAtt!allow_read), 0, rsAtt!allow_read)
                    cmd.SetInt32 9, IIf(IsNull(rsAtt!allow_add), 0, rsAtt!allow_add)
                    cmd.SetInt32 10, IIf(IsNull(rsAtt!allow_edit), 0, rsAtt!allow_edit)
                    cmd.SetInt32 11, IIf(IsNull(rsAtt!allow_delete), 0, rsAtt!allow_delete)
                    cmd.SetInt32 12, IIf(IsNull(rsAtt!allow_post), 0, rsAtt!allow_post)
                    cmd.SetInt32 13, IIf(IsNull(rsAtt!allow_print), 0, rsAtt!allow_print)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
'    v_company = "Total Data Perusahaan = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataUserAccess()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS t_user_access_level (" & _
                        "level_code TEXT(30) NOT NULL,access_level_code INTEGER NOT NULL," & _
                        "level_name TEXT(50) DEFAULT NULL,allow_access INTEGER DEFAULT NULL," & _
                        "PRIMARY KEY(level_code,access_level_code)); "
                        
        strsql = "INSERT INTO t_user_access_level " & _
                    "(level_code,access_level_code,level_name,allow_access) " & _
                    "Values " & _
                    "(?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT level_code,access_level_code,level_name,allow_access " & _
                 "FROM t_user_access_level"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM t_user_access_level")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!level_code
                    cmd.SetInt32 2, IIf(IsNull(rsAtt!access_level_code), 0, rsAtt!access_level_code)
                    cmd.SetText 3, IIf(IsNull(rsAtt!level_name), "", rsAtt!level_name)
                    cmd.SetInt32 4, IIf(IsNull(rsAtt!allow_access), 0, rsAtt!allow_access)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
'    v_company = "Total Data Perusahaan = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataShift()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS m_shift (" & _
                        "shift_code TEXT(20) NOT NULL,shift_name TEXT(30) DEFAULT NULL,start_time TEXT(30) DEFAULT NULL," & _
                        "end_time TEXT(30) DEFAULT NULL,flag_day_over INTEGER DEFAULT NULL,flag_tolerance INTEGER DEFAULT NULL," & _
                        "start_time_tolerance TEXT(30) DEFAULT NULL,end_time_tolerance TEXT(30) DEFAULT NULL," & _
                        "flag_shift INTEGER DEFAULT NULL,min_break_in TEXT(30) DEFAULT NULL,max_break_out TEXT(30) DEFAULT NULL," & _
                        "break_interval_minute INTEGER DEFAULT NULL,flag_moving INTEGER DEFAULT NULL,moving_number INTEGER DEFAULT NULL," & _
                        "company_code TEXT(30) DEFAULT NULL,PRIMARY KEY(shift_code,company_code)); "
                        
        strsql = "INSERT INTO m_shift " & _
                    "(shift_code,shift_name,start_time,end_time,flag_day_over,flag_tolerance," & _
                        "start_time_tolerance,end_time_tolerance,flag_shift," & _
                        "min_break_in,max_break_out,break_interval_minute," & _
                        "flag_moving,moving_number,company_code) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT * FROM m_shift " _
                & "WHERE company_code = '" & TDBCombo_company.Text & "'"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM m_shift")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!shift_code
                    cmd.SetText 2, IIf(IsNull(rsAtt!shift_name) = True, "", rsAtt!shift_name)
                    cmd.SetDate 3, IIf(IsNull(rsAtt!start_time) = True, 0, rsAtt!start_time)
                    cmd.SetDate 4, IIf(IsNull(rsAtt!end_time) = True, 0, rsAtt!end_time)
                    cmd.SetInt32 5, IIf(IsNull(rsAtt!flag_day_over) = True, 0, rsAtt!flag_day_over)
                    cmd.SetInt32 6, IIf(IsNull(rsAtt!flag_tolerance) = True, 0, rsAtt!flag_tolerance)
                    cmd.SetDate 7, IIf(IsNull(rsAtt!start_time_tolerance) = True, 0, rsAtt!start_time_tolerance)
                    cmd.SetDate 8, IIf(IsNull(rsAtt!end_time_tolerance) = True, 0, rsAtt!end_time_tolerance)
                    cmd.SetInt32 9, IIf(IsNull(rsAtt!flag_shift) = True, 0, rsAtt!flag_shift)
                    cmd.SetDate 10, IIf(IsNull(rsAtt!min_break_in) = True, 0, rsAtt!min_break_in)
                    cmd.SetDate 11, IIf(IsNull(rsAtt!max_break_out) = True, 0, rsAtt!max_break_out)
                    cmd.SetInt32 12, IIf(IsNull(rsAtt!break_interval_minute) = True, 0, rsAtt!break_interval_minute)
                    cmd.SetInt32 13, IIf(IsNull(rsAtt!flag_moving) = True, 0, rsAtt!flag_moving)
                    cmd.SetInt32 14, IIf(IsNull(rsAtt!moving_number) = True, 0, rsAtt!moving_number)
                    cmd.SetText 15, rsAtt!COMPANY_CODE
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    v_shift = "Total Data Shift = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataSalary()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS h_salary (" & _
                        "month TEXT(20) NOT NULL,employee_code TEXT(30) NOT NULL,salary_code TEXT(20) DEFAULT NULL," & _
                        "salary_name TEXT(50) DEFAULT NULL,date_from TEXT(30) DEFAULT NULL,date_to TEXT(30) DEFAULT NULL," & _
                        "flag_main_salary TEXT(30) DEFAULT NULL,flag_sign INTEGER DEFAULT NULL," & _
                        "flag_detail INTEGER DEFAULT NULL,flag_use_formula INTEGER DEFAULT NULL,formula_salary_code TEXT(30) DEFAULT NULL," & _
                        "flag_ptkp INTEGER DEFAULT NULL,ptkp_salary_code TEXT(30) DEFAULT NULL,flag_pkp INTEGER DEFAULT NULL," & _
                        "flag_pph21 INTEGER DEFAULT NULL,pph21_number INTEGER DEFAULT NULL,flag_tax INTEGER DEFAULT NULL," & _
                        "tax_salary_code TEXT(30) DEFAULT NULL,flag_type TEXT(10) DEFAULT NULL,flag_visible INTEGER DEFAULT NULL," & _
                        "salary_value DECIMAL(15,2) DEFAULT NULL,description TEXT(50) DEFAULT NULL,company_code TEXT(30) DEFAULT NULL," & _
                        "PRIMARY KEY(month,employee_code,salary_code)); "
                        
        strsql = "INSERT INTO h_salary " & _
                    "(month,employee_code,salary_code,salary_name,date_from,date_to," & _
                    "flag_main_salary,flag_sign,flag_detail,flag_use_formula,formula_salary_code," & _
                    "flag_ptkp,ptkp_salary_code,flag_pkp,flag_pph21,pph21_number,flag_tax," & _
                    "tax_salary_code,flag_type,flag_visible,salary_value,description,company_code) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT a.* FROM h_salary a join m_employee b on a.employee_code = b.employee_code " _
                & "WHERE month = '" & Format(DTPicker_periode_from.Value, "yyyy-MM") & "' AND " _
                & "CASE WHEN '" & Check2.Value & "' = 1 then a.company_code = '" & TDBCombo_company.Text & "' AND b.department_code = '" & TDBCombo_department.Text & "' " _
                & "ELSE a.company_code = '" & TDBCombo_company.Text & "' END"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM h_salary")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!month
                    cmd.SetText 2, rsAtt!employee_code
                    cmd.SetText 3, rsAtt!salary_code
                    cmd.SetText 4, IIf(IsNull(rsAtt!salary_name) = True, "", rsAtt!salary_name)
                    cmd.SetText 5, IIf(IsNull(rsAtt!date_from) = True, "", rsAtt!date_from)
                    cmd.SetText 6, IIf(IsNull(rsAtt!date_to) = True, "", rsAtt!date_to)
                    cmd.SetText 7, IIf(IsNull(rsAtt!flag_main_salary) = True, "", rsAtt!flag_main_salary)
                    cmd.SetInt32 8, IIf(IsNull(rsAtt!flag_sign) = True, 0, rsAtt!flag_sign)
                    cmd.SetInt32 9, IIf(IsNull(rsAtt!flag_detail) = True, 0, rsAtt!flag_detail)
                    cmd.SetInt32 10, IIf(IsNull(rsAtt!flag_use_formula) = True, 0, rsAtt!flag_use_formula)
                    cmd.SetText 11, IIf(IsNull(rsAtt!formula_salary_code) = True, "", rsAtt!formula_salary_code)
                    cmd.SetInt32 12, IIf(IsNull(rsAtt!flag_ptkp) = True, 0, rsAtt!flag_ptkp)
                    cmd.SetText 13, IIf(IsNull(rsAtt!ptkp_salary_code) = True, "", rsAtt!ptkp_salary_code)
                    cmd.SetInt32 14, IIf(IsNull(rsAtt!flag_pkp) = True, 0, rsAtt!flag_pkp)
                    cmd.SetInt32 15, IIf(IsNull(rsAtt!flag_pph21) = True, 0, rsAtt!flag_pph21)
                    cmd.SetInt32 16, IIf(IsNull(rsAtt!pph21_number) = True, 0, rsAtt!pph21_number)
                    cmd.SetInt32 17, IIf(IsNull(rsAtt!flag_tax) = True, 0, rsAtt!flag_tax)
                    cmd.SetText 18, IIf(IsNull(rsAtt!tax_salary_code) = True, "", rsAtt!tax_salary_code)
                    cmd.SetText 19, IIf(IsNull(rsAtt!flag_type) = True, "", rsAtt!flag_type)
                    cmd.SetInt32 20, IIf(IsNull(rsAtt!flag_visible) = True, 0, rsAtt!flag_visible)
                    cmd.SetDouble 21, IIf(IsNull(rsAtt!salary_value) = True, 0, rsAtt!salary_value)
                    cmd.SetText 22, IIf(IsNull(rsAtt!Description) = True, "", rsAtt!Description)
                    cmd.SetText 23, Trim(TDBCombo_company.Text)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    MsgBox "Transfer Data Berhasil. Total Data Transfer = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub exportDataHDSalary()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS h_d_salary (" & _
                        "month TEXT(20) NOT NULL,periode_from TEXT(30) DEFAULT NULL," & _
                        "periode_to TEXT(20) DEFAULT NULL,company_code TEXT(30) NOT NULL," & _
                        "company_name TEXT(100) DEFAULT NULL,PRIMARY KEY(month,company_code)); "
                        
        strsql = "INSERT INTO h_d_salary " & _
                    "(month,periode_from,periode_to,company_code,company_name) " & _
                    "Values " & _
                    "(?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT month,periode_from,periode_to,company_code,company_name " & _
                 "FROM h_d_salary"
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If

        Label2.Visible = True
        ProgressBar1.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM h_d_salary")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetDate 1, rsAtt!month
                    cmd.SetDate 2, IIf(IsNull(rsAtt!periode_from), 0, rsAtt!periode_from)
                    cmd.SetDate 3, IIf(IsNull(rsAtt!periode_to), 0, rsAtt!periode_to)
                    cmd.SetText 4, IIf(IsNull(rsAtt!COMPANY_CODE), "", rsAtt!COMPANY_CODE)
                    cmd.SetText 5, IIf(IsNull(rsAtt!company_name), "", rsAtt!company_name)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            rsAtt.Close
            conLite.CommitTrans
        'End If
    
    ProgressBar1.Visible = False
    Label2.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
'    v_department = "Total Data Departemen / Area = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub importDataEmployee()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

'On Error GoTo err_capture

    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_employee") Then
        MsgBox "Tidak Bisa MEnemukan Data Pada FIle Yang Dipilih... Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM m_employee"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
        v_company_code = rss!COMPANY_CODE
            
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Untuk Perusahaan Ini Tidak Valid, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
        
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_employee where company_code = '" & TDBCombo_company.Text & "'", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If

    For aa = 0 To rss.RecordCount - 1
'    rs.MoveFirst
'    Do While Not rs.EOF
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
        
        If Not check_exist_new_employee(rss!employee_code) Then
            i = i + 1
                
            With rs
'                If Not check_exist_new_employee(rss!employee_code) Then
                .AddNew
'                End If
                                
                .Fields("employee_code").Value = rss!employee_code
                .Fields("nik").Value = rss!nik
                '-----------------------------------------------------------------------------
                .Fields("employee_name").Value = rss!EMPLOYEE_NAME
                .Fields("employee_nick_name").Value = rss!employee_nick_name
                
                .Fields("company_code").Value = rss!COMPANY_CODE
                .Fields("company_name").Value = rss!company_name
                .Fields("company_name").Value = rss!company_name
                .Fields("department_code").Value = rss!DEPARTMENT_CODE
                .Fields("department_name").Value = rss!department_name
                .Fields("division_code").Value = Replace(rss!division_code, "'", "''")
                .Fields("division_name").Value = rss!division_name
                
                .Fields("date_of_birth").Value = IIf(IsNull(rss!date_of_birth) = True, "0000-00-00 00:00:00", rss!date_of_birth)
                .Fields("place_of_birth").Value = rss!place_of_birth
                .Fields("sex").Value = rss!sex
                .Fields("religion").Value = rss!religion
                .Fields("marital_status").Value = rss!marital_status
                .Fields("number_of_children").Value = rss!number_of_children
                .Fields("address").Value = rss!Address
                .Fields("email").Value = rss!email
                .Fields("npwp").Value = rss!npwp
                
                .Fields("phone_number").Value = rss!phone_number
                .Fields("bank_account").Value = rss!bank_account
                .Fields("last_education_code").Value = rss!last_education_code
            
                .Fields("last_education_name").Value = rss!last_education_name
                '.Fields("last_education_pass").Value = Format(isnullrss!last_education_pass, "yyyy-MM-dd HH:mm:ss")
                .Fields("last_education_pass").Value = IIf(rss!last_education_pass = "", "00:00:00", Format(rss!last_education_pass, "yyyy-mm-dd HH:mm:ss"))
                .Fields("last_employment_name").Value = rss!last_employment_name
                .Fields("last_employment_date").Value = IIf(rss!last_employment_date = "", "00:00:00", Format(rss!last_education_pass, "yyyy-mm-dd HH:mm:ss"))
                .Fields("last_employment_title").Value = rss!last_employment_title
                .Fields("start_working").Value = rss!start_working
                .Fields("date_of_appointment").Value = IIf(rss!date_of_appointment = "", "00:00:00", Format(rss!last_education_pass, "yyyy-mm-dd HH:mm:ss"))
                
                .Fields("title_code").Value = rss!title_code
                .Fields("title_name").Value = rss!title_name
                
                .Fields("level1").Value = rss!level1
                .Fields("level2").Value = rss!level2
                
                .Fields("flag_shiftable").Value = rss!flag_shiftable
                .Fields("flag_active").Value = rss!flag_active
                
                .Fields("fathers_name").Value = rss!fathers_name
                .Fields("mothers_name").Value = rss!mothers_name
                .Fields("child_number").Value = rss!child_number
                .Fields("child_number_from").Value = rss!child_number_from
                
                
                .Fields("description").Value = rss!Description
                .Fields("end_working").Value = IIf(rss!end_working = "", "00:00:00", Format(rss!last_education_pass, "yyyy-mm-dd HH:mm:ss"))
                .Fields("reason").Value = rss!reason
                
                .Fields("all_in").Value = rss!all_in
                .Fields("no_jamsostek").Value = rss!no_jamsostek
                .Fields("level_code").Value = rss!level_code
                .Fields("pertanggungan_pajak").Value = rss!pertanggungan_pajak
                .Fields("grade").Value = rss!grade
                .Fields("status_employee").Value = rss!status_employee
'                .Fields("area_code").Value = rss!area_code
                .Fields("prev_location").Value = rss!prev_location
                .Fields("curr_location").Value = rss!curr_location
                .Fields("start_mc").Value = IIf(rss!start_mc = "", "00:00:00", Format(rss!start_mc, "yyyy-MM-dd"))
                .Fields("end_mc").Value = IIf(rss!end_mc = "", "00:00:00", Format(rss!end_mc, "yyyy-MM-dd"))
                .Fields("account_name").Value = rss!account_name
                .Fields("seq_no").Value = rss!seq_no
                
                .Update
            End With
        Else
            strsql = "UPDATE m_employee SET employee_code = '" & rss!employee_code & "',employee_name = '" & Replace(rss!EMPLOYEE_NAME, "'", "''") & "'," _
                & "employee_nick_name = '" & Replace(rss!employee_nick_name, "'", "''") & "',company_code = '" & rss!COMPANY_CODE & "'," _
                & "company_name = '" & rss!company_name & "',department_code = '" & rss!DEPARTMENT_CODE & "'," _
                & "department_name = '" & rss!department_name & "',division_code = '" & Replace(rss!division_code, "'", "''") & "'," _
                & "division_name = '" & rss!division_name & "',date_of_birth = '" & IIf(IsNull(rss!date_of_birth) Or rss!date_of_birth = "", "0000-00-00 00:00:00", Format(rss!date_of_birth, "yyyy-MM-dd hh:nn:ss")) & "'," _
                & "place_of_birth = '" & rss!place_of_birth & "',sex = '" & rss!sex & "'," _
                & "religion = '" & rss!religion & "',marital_status = '" & rss!marital_status & "'," _
                & "number_of_children = '" & rss!number_of_children & "',address = '" & rss!Address & "'," _
                & "email = '" & rss!email & "',npwp = '" & rss!npwp & "'," _
                & "phone_number = '" & rss!phone_number & "',bank_account = '" & rss!bank_account & "'," _
                & "last_education_code = '" & rss!last_education_code & "',last_education_name = '" & rss!last_education_name & "'," _
                & "last_education_pass = '" & IIf(rss!last_education_pass = "", "00:00:00", Format(rss!last_education_pass, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "last_employment_name = '" & rss!last_employment_name & "',last_employment_date = '" & IIf(rss!last_employment_date = "", "00:00:00", Format(rss!last_education_pass, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "last_employment_title = '" & rss!last_employment_title & "',start_working = '" & IIf(IsNull(rss!start_working) Or rss!start_working = "", "0000-00-00 00:00:00", Format(rss!start_working, "yyyy-MM-dd hh:nn:ss")) & "'," _
                & "date_of_appointment = '" & IIf(rss!date_of_appointment = "", "00:00:00", Format(rss!last_education_pass, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "title_code = '" & rss!title_code & "',title_name = '" & rss!title_name & "'," _
                & "level1 = '" & rss!level1 & "',level2 = '" & rss!level2 & "',flag_shiftable = '" & rss!flag_shiftable & "'," _
                & "flag_active = '" & rss!flag_active & "',fathers_name = '" & rss!fathers_name & "',mothers_name = '" & rss!mothers_name & "',child_number = '" & rss!child_number & "'," _
                & "child_number_from = '" & rss!child_number_from & "',description = '" & rss!Description & "'," _
                & "end_working = '" & IIf(rss!end_working = "", "00:00:00", Format(rss!last_education_pass, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "reason = '" & rss!reason & "',all_in = '" & rss!all_in & "',no_jamsostek = '" & rss!no_jamsostek & "'," _
                & "level_code = '" & rss!level_code & "',pertanggungan_pajak = '" & rss!pertanggungan_pajak & "',grade = '" & rss!grade & "',status_employee = '" & rss!status_employee & "'," _
                & "prev_location = '" & rss!prev_location & "',curr_location = '" & rss!curr_location & "',nik = '" & rss!nik & "'," _
                & "start_mc = '" & IIf(rss!start_mc = "", "00:00:00", Format(rss!start_mc, "yyyy-MM-dd")) & "',end_mc = '" & IIf(rss!end_mc = "", "00:00:00", Format(rss!end_mc, "yyyy-MM-dd")) & "',account_name = '" & Replace(rss!account_name, "'", "''") & "',seq_no = '" & rss!seq_no & "' " _
                & "WHERE employee_code = '" & rss!employee_code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
        
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_employee = "Total Data Karyawan = " & i & " !"
    Exit Sub
'Exit Sub
'err_capture:
'rs.CancelBatch adAffectCurrent: rs.Close: CnG.RollbackTrans
End Sub

Private Sub importDataCompany()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_company") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM m_company"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_company", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_company(rss!COMPANY_CODE) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("company_code").Value = rss!COMPANY_CODE
                .Fields("company_name").Value = rss!company_name
                .Fields("address").Value = rss!Address
                .Fields("postal_code").Value = rss!postal_code
                .Fields("city_name").Value = rss!city_name
                .Fields("phone_number").Value = rss!phone_number
                .Fields("fax_number").Value = rss!fax_number
                .Fields("web_address").Value = rss!web_address
                .Fields("email_address").Value = rss!email_address
                .Fields("npwp").Value = rss!npwp
                .Fields("pimpinan").Value = rss!pimpinan
                .Fields("pimpinan_npwp").Value = rss!pimpinan_npwp
                .Fields("npp").Value = rss!npp
                .Fields("hari_kerja").Value = rss!hari_kerja
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE m_company SET company_code = '" & rss!COMPANY_CODE & "'," _
                & "company_name = '" & rss!company_name & "'," _
                & "address = '" & rss!Address & "'," _
                & "postal_code = '" & rss!postal_code & "'," _
                & "city_name = '" & rss!city_name & "'," _
                & "phone_number = '" & rss!phone_number & "'," _
                & "fax_number = '" & rss!fax_number & "'," _
                & "web_address = '" & rss!web_address & "'," _
                & "email_address = '" & rss!email_address & "'," _
                & "npwp = '" & rss!npwp & "'," _
                & "pimpinan = '" & rss!pimpinan & "'," _
                & "pimpinan_npwp = '" & rss!pimpinan_npwp & "'," _
                & "npp = '" & rss!npp & "'," _
                & "hari_kerja = '" & rss!hari_kerja & "' " _
                & "WHERE company_code = '" & rss!COMPANY_CODE & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_company = "Total Data Perusahaan = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataDepartment()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_department") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM m_department"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_department", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_department(rss!DEPARTMENT_CODE, rss!COMPANY_CODE) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("company_code").Value = rss!COMPANY_CODE
                .Fields("department_code").Value = rss!DEPARTMENT_CODE
                .Fields("department_name").Value = rss!department_name
                .Fields("description").Value = rss!Description
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE m_department SET company_code = '" & rss!COMPANY_CODE & "'," _
                & "department_code = '" & rss!DEPARTMENT_CODE & "'," _
                & "department_name = '" & rss!department_name & "'," _
                & "description = '" & rss!Description & "' " _
                & "WHERE company_code = '" & rss!COMPANY_CODE & "' AND " _
                & "department_code = '" & rss!DEPARTMENT_CODE & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_department = "Total Data Departement / Area = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataDivision()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_division") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM m_division"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_division", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_division(rss!division_code, rss!DEPARTMENT_CODE, rss!COMPANY_CODE) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("company_code").Value = rss!COMPANY_CODE
                .Fields("department_code").Value = rss!DEPARTMENT_CODE
                .Fields("department_name").Value = rss!department_name
                .Fields("division_code").Value = Replace(rss!division_code, "'", "''")
                .Fields("division_name").Value = rss!division_name
                .Fields("description").Value = rss!Description
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE m_division SET company_code = '" & rss!COMPANY_CODE & "'," _
                & "department_code = '" & rss!DEPARTMENT_CODE & "'," _
                & "department_name = '" & rss!department_name & "'," _
                & "division_code = '" & Replace(rss!division_code, "'", "''") & "'," _
                & "division_name = '" & rss!division_name & "'," _
                & "description = '" & rss!Description & "' " _
                & "WHERE company_code = '" & rss!COMPANY_CODE & "' AND " _
                & "department_code = '" & rss!DEPARTMENT_CODE & "' AND " _
                & "division_code = '" & rss!division_code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_division = "Total Data Divisi = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataLocation()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_location") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM m_location"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_location", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_location(rss!location_code, rss!COMPANY_CODE) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("company_code").Value = rss!COMPANY_CODE
                .Fields("location_code").Value = rss!location_code
                .Fields("location_name").Value = rss!location_name
                .Fields("description").Value = rss!Description
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE m_location SET company_code = '" & rss!COMPANY_CODE & "'," _
                & "location_code = '" & rss!location_code & "'," _
                & "location_name = '" & rss!location_name & "'," _
                & "description = '" & rss!Description & "' " _
                & "WHERE company_code = '" & rss!COMPANY_CODE & "' AND " _
                & "location_code = '" & rss!location_code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_location = "Total Data Lokasi = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataBank()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_bank") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM m_bank"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
'        If rss.RecordCount > 0 Then
'            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
'                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
'                Exit Sub
'            End If
'        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_bank", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_bank(rss!bank_code) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("bank_code").Value = rss!bank_code
                .Fields("bank_name").Value = rss!bank_name
                .Fields("description").Value = rss!Description
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE m_bank SET " _
                & "bank_code = '" & rss!bank_code & "'," _
                & "bank_name = '" & rss!bank_name & "'," _
                & "description = '" & rss!Description & "' " _
                & "WHERE bank_code = '" & rss!bank_code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_bank = "Total Data Bank = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataTitle()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_title") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM m_title"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_title", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_title(rss!title_code, rss!COMPANY_CODE) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("title_code").Value = rss!title_code
                .Fields("title_name").Value = rss!title_name
                .Fields("default_shiftable").Value = IIf(IsNull(rss!default_shiftable), 0, rss!default_shiftable)
                .Fields("level").Value = rss!Level
                .Fields("description").Value = rss!Description
                .Fields("salary").Value = rss!salary
                .Fields("company_code").Value = rss!COMPANY_CODE
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE m_title SET title_code = '" & rss!title_code & "'," _
                & "title_name = '" & rss!title_name & "'," _
                & "default_shiftable = '" & IIf(IsNull(rss!default_shiftable), 0, rss!default_shiftable) & "'," _
                & "level = '" & rss!Level & "'," _
                & "description = '" & rss!Description & "'," _
                & "salary = '" & rss!salary & "'," _
                & "company_code = '" & rss!COMPANY_CODE & "' " _
                & "WHERE company_code = '" & rss!COMPANY_CODE & "' AND " _
                & "title_code = '" & rss!title_code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_title = "Total Data Jabatan = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataLevel()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_akses_level_group") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM m_akses_level_group"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
'        If rss.RecordCount > 0 Then
'            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
'                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
'                Exit Sub
'            End If
'        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_akses_level_group", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_level(rss!Code) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("code").Value = rss!Code
                .Fields("date_entry").Value = IIf(rss!date_entry = "", "00:00:00", Format(rss!date_entry, "yyyy-mm-dd HH:mm:ss"))
                .Fields("date_edit").Value = IIf(rss!date_edit = "", "00:00:00", Format(rss!date_edit, "yyyy-mm-dd HH:mm:ss"))
                .Fields("remark").Value = rss!remark
                .Fields("user_entry").Value = rss!user_entry
                .Fields("user_edit").Value = rss!user_edit
                .Fields("name").Value = rss!name
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE m_akses_level_group SET code = '" & rss!Code & "'," _
                & "date_entry = '" & IIf(rss!date_entry = "", "00:00:00", Format(rss!date_entry, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "date_edit = '" & IIf(rss!date_edit = "", "00:00:00", Format(rss!date_edit, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "remark = '" & rss!remark & "'," _
                & "user_entry = '" & rss!user_entry & "'," _
                & "user_edit = '" & rss!user_edit & "'," _
                & "name = '" & rss!name & "' " _
                & "WHERE code = '" & rss!Code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_level = "Total Data Level = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataAttendance()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("h_attendance") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM h_attendance"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from h_attendance", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
    
    For aa = 0 To rss.RecordCount - 1
        
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_attendance(rss!employee_code, rss!att_date) Then
            i = i + 1
                
            With rs
                .AddNew
                '-----------------------
                .Fields("employee_code").Value = rss!employee_code
                .Fields("att_date").Value = rss!att_date
                .Fields("ip_address").Value = rss!ip_address
                .Fields("enrollnumber").Value = rss!enrollnumber
                .Fields("shift_number").Value = rss!shift_number
                .Fields("shift_code").Value = rss!shift_code
                .Fields("start_time").Value = IIf(rss!start_time = "", "00:00:00", Format(rss!start_time, "yyyy-mm-dd HH:mm:ss"))
                .Fields("end_time").Value = IIf(rss!end_time = "", "00:00:00", Format(rss!end_time, "yyyy-mm-dd HH:mm:ss"))
                .Fields("time_in").Value = IIf(rss!time_in = "", "00:00:00", Format(rss!time_in, "yyyy-mm-dd HH:mm:ss"))
                .Fields("time_out").Value = IIf(rss!time_out = "", "00:00:00", Format(rss!time_out, "yyyy-mm-dd HH:mm:ss"))
                .Fields("time_in_diff").Value = IIf(rss!time_in_diff = "", "00:00:00", Format(rss!time_in_diff, "yyyy-mm-dd HH:mm:ss"))
                .Fields("time_out_diff").Value = IIf(rss!time_out_diff = "", "00:00:00", Format(rss!time_out_diff, "yyyy-mm-dd HH:mm:ss"))
                .Fields("flag_io").Value = rss!flag_io
                .Fields("flag_present").Value = rss!flag_present
                .Fields("flag_duty").Value = rss!flag_duty
                .Fields("flag_late").Value = rss!flag_late
                .Fields("flag_early").Value = rss!flag_early
                .Fields("absent_status").Value = rss!absent_status
                .Fields("time_in_break").Value = IIf(rss!time_in_break = "", "00:00:00", Format(rss!time_in_break, "yyyy-mm-dd HH:mm:ss"))
                .Fields("time_in_break_diff").Value = IIf(rss!time_in_break_diff = "", "00:00:00", Format(rss!time_in_break_diff, "yyyy-mm-dd HH:mm:ss"))
                .Fields("flag_in_break_early").Value = rss!flag_in_break_early
                .Fields("time_out_break").Value = IIf(rss!time_out_break = "", "00:00:00", Format(rss!time_out_break, "yyyy-mm-dd HH:mm:ss"))
                .Fields("time_out_break_diff").Value = IIf(rss!time_out_break_diff = "", "00:00:00", Format(rss!time_out_break_diff, "yyyy-mm-dd HH:mm:ss"))
                .Fields("flag_out_break_late").Value = rss!flag_out_break_late
                .Fields("break_interval").Value = IIf(rss!break_interval = "", "00:00:00", Format(rss!break_interval, "yyyy-mm-dd HH:mm:ss"))
                .Fields("break_ot").Value = IIf(rss!break_ot = "", "00:00:00", Format(rss!break_ot, "yyyy-mm-dd HH:mm:ss"))
                .Fields("Description").Value = rss!Description
                .Fields("entry_date").Value = IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss"))
                .Fields("editdate").Value = IIf(rss!edit_date = "", "00:00:00", Format(rss!edit_date, "yyyy-mm-dd HH:mm:ss"))
                .Fields("userinput").Value = rss!userInput
                .Fields("useredit").Value = rss!userEdit
                .Fields("total_ot").Value = rss!total_OT
                .Fields("15jam").Value = rss!satu_jam
                .Fields("2jam").Value = rss!dua_jam
                .Fields("3jam").Value = rss!tiga_jam
                .Fields("4jam").Value = rss!empat_jam
                .Fields("holiday").Value = rss!holiday
                .Fields("tot_overtime").Value = rss!tot_overtime
                .Fields("flag_meal").Value = rss!flag_meal
                .Fields("flag_transport").Value = rss!flag_transport
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE h_attendance SET employee_code = '" & rss!employee_code & "',att_date = '" & rss!att_date & "'," _
                & "ip_address = '" & rss!ip_address & "',enrollnumber = '" & rss!enrollnumber & "',shift_number = '" & rss!shift_number & "',shift_code = '" & rss!shift_code & "'," _
                & "start_time = '" & IIf(rss!start_time = "", "00:00:00", Format(rss!start_time, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "end_time = '" & IIf(rss!end_time = "", "00:00:00", Format(rss!end_time, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "time_in = '" & IIf(rss!time_in = "", "00:00:00", Format(rss!time_in, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "time_out = '" & IIf(rss!time_out = "", "00:00:00", Format(rss!time_out, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "time_in_diff = '" & IIf(rss!time_in_diff = "", "00:00:00", Format(rss!time_in_diff, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "time_out_diff = '" & IIf(rss!time_out_diff = "", "00:00:00", Format(rss!time_out_diff, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "flag_io = '" & rss!flag_io & "',flag_present = '" & rss!flag_present & "',flag_duty = '" & rss!flag_duty & "',flag_late = '" & rss!flag_late & "'," _
                & "flag_early = '" & rss!flag_early & "',absent_status = '" & IIf(rss!absent_status = "" Or IsNull(rss!absent_status), 0, rss!absent_status) & "'," _
                & "time_in_break = '" & IIf(rss!time_in_break = "", "00:00:00", Format(rss!time_in_break, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "time_in_break_diff = '" & IIf(rss!time_in_break_diff = "", "00:00:00", Format(rss!time_in_break_diff, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "flag_in_break_early = '" & rss!flag_in_break_early & "',time_out_break = '" & IIf(rss!time_out_break = "", "00:00:00", Format(rss!time_out_break, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "time_out_break_diff = '" & IIf(rss!time_out_break_diff = "", "00:00:00", Format(rss!time_out_break_diff, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "flag_out_break_late = '" & rss!flag_out_break_late & "',break_interval = '" & IIf(rss!break_interval = "", "00:00:00", Format(rss!break_interval, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "break_ot = '" & IIf(rss!break_ot = "", "00:00:00", Format(rss!break_ot, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "Description = '" & rss!Description & "',entry_date = '" & IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "editdate = '" & IIf(rss!edit_date = "", "00:00:00", Format(rss!edit_date, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "userinput = '" & rss!userInput & "',useredit = '" & rss!userEdit & "'," _
                & "total_ot = '" & rss!total_OT & "',15jam = '" & rss!satu_jam & "'," _
                & "2jam = '" & rss!dua_jam & "',3jam = '" & rss!tiga_jam & "'," _
                & "4jam = '" & rss!empat_jam & "',holiday = '" & rss!holiday & "'," _
                & "tot_overtime = '" & rss!tot_overtime & "',flag_meal = '" & rss!flag_meal & "',flag_transport = '" & rss!flag_transport & "' " _
                & "WHERE employee_code = '" & rss!employee_code & "' AND " _
                & "date(att_date) = '" & Format(rss!att_date, "yyyy-MM-dd") & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_attendance = "Total Data Kehadiran = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataAbsent()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("t_absent") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM t_absent"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from t_absent", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_absent(rss!absent_number) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("absent_number").Value = rss!absent_number
                .Fields("employee_code").Value = rss!employee_code
                .Fields("absent_date_from").Value = IIf(rss!absent_date_from = "", "00:00:00", Format(rss!absent_date_from, "yyyy-mm-dd HH:mm:ss"))
                .Fields("flag_date_to").Value = rss!flag_date_to
                .Fields("absent_date_to").Value = IIf(rss!absent_date_to = "", "00:00:00", Format(rss!absent_date_to, "yyyy-mm-dd HH:mm:ss"))
                .Fields("absent_status").Value = rss!absent_status
                .Fields("description").Value = IIf(IsNull(rss!Description), "", rss!Description)
                .Fields("entry_date").Value = IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss"))
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE t_absent SET absent_number = '" & rss!absent_number & "'," _
                & "employee_code = '" & rss!employee_code & "'," _
                & "absent_date_from = '" & IIf(rss!absent_date_from = "", "00:00:00", Format(rss!absent_date_from, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "flag_date_to = '" & rss!flag_date_to & "'," _
                & "absent_date_to = '" & IIf(rss!absent_date_to = "", "00:00:00", Format(rss!absent_date_to, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "absent_status = '" & rss!absent_status & "'," _
                & "description = '" & IIf(IsNull(rss!Description), "", rss!Description) & "'," _
                & "entry_date = '" & IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss")) & "' " _
                & "WHERE absent_number = '" & rss!absent_number & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_absent = "Total Data Absensi = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataDuty()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("t_duty") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM t_duty"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from t_duty", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_duty(rss!duty_number) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("duty_number").Value = rss!duty_number
                .Fields("employee_code").Value = rss!employee_code
                .Fields("duty_date_from").Value = IIf(rss!duty_date_from = "", "00:00:00", Format(rss!duty_date_from, "yyyy-mm-dd HH:mm:ss"))
                .Fields("flag_date_to").Value = rss!flag_date_to
                .Fields("duty_date_to").Value = IIf(rss!duty_date_to = "", "00:00:00", Format(rss!duty_date_to, "yyyy-mm-dd HH:mm:ss"))
                .Fields("description").Value = IIf(IsNull(rss!Description), "", rss!Description)
                .Fields("entry_date").Value = IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss"))
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE t_duty SET duty_number = '" & rss!duty_number & "'," _
                & "employee_code = '" & rss!employee_code & "'," _
                & "duty_date_from = '" & IIf(rss!duty_date_from = "", "00:00:00", Format(rss!duty_date_from, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "flag_date_to = '" & rss!flag_date_to & "'," _
                & "duty_date_to = '" & IIf(rss!duty_date_to = "", "00:00:00", Format(rss!duty_date_to, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "description = '" & IIf(IsNull(rss!Description), "", rss!Description) & "'," _
                & "entry_date = '" & IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss")) & "' " _
                & "WHERE duty_number = '" & rss!duty_number & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_duty = "Total Data Tugas Dinas = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataLeave()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("t_leave") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM t_leave"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from t_leave", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_leave(rss!leave_number) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("leave_number").Value = rss!leave_number
                .Fields("employee_code").Value = rss!employee_code
                .Fields("leave_type").Value = rss!leave_type
                .Fields("flag_doctor_order").Value = rss!flag_doctor_order
                .Fields("leave_date_from").Value = IIf(rss!leave_date_from = "", "00:00:00", Format(rss!leave_date_from, "yyyy-mm-dd HH:mm:ss"))
                .Fields("flag_date_to").Value = rss!flag_date_to
                .Fields("leave_date_to").Value = IIf(rss!leave_date_to = "", "00:00:00", Format(rss!leave_date_to, "yyyy-mm-dd HH:mm:ss"))
                .Fields("description").Value = IIf(IsNull(rss!Description), "", rss!Description)
                .Fields("entry_date").Value = IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss"))
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE t_leave SET leave_number = '" & rss!leave_number & "'," _
                & "employee_code = '" & rss!employee_code & "'," _
                & "leave_type = '" & rss!leave_type & "'," _
                & "flag_doctor_order = '" & rss!flag_doctor_order & "'," _
                & "leave_date_from = '" & IIf(rss!leave_date_from = "", "00:00:00", Format(rss!leave_date_from, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "flag_date_to = '" & rss!flag_date_to & "'," _
                & "leave_date_to = '" & IIf(rss!leave_date_to = "", "00:00:00", Format(rss!leave_date_to, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "description = '" & IIf(IsNull(rss!Description), "", rss!Description) & "'," _
                & "entry_date = '" & IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss")) & "' " _
                & "WHERE leave_number = '" & rss!leave_number & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_leave = "Total Data Cuti = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataGeneralLeave()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("t_general_leave") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM t_general_leave"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
'        If rss.RecordCount > 0 Then
'            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
'                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
'                Exit Sub
'            End If
'        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from t_general_leave", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_general_leave(rss!general_leave_number) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("general_leave_number").Value = rss!general_leave_number
                .Fields("general_leave_date").Value = IIf(rss!general_leave_date = "", "00:00:00", Format(rss!general_leave_date, "yyyy-mm-dd HH:mm:ss"))
                .Fields("description").Value = IIf(IsNull(rss!Description), "", rss!Description)
                .Fields("entry_date").Value = IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss"))
                .Fields("company_code").Value = rss!COMPANY_CODE
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE t_general_leave SET general_leave_number = '" & rss!general_leave_number & "'," _
                & "general_leave_date = '" & IIf(rss!general_leave_date = "", "00:00:00", Format(rss!general_leave_date, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "description = '" & IIf(IsNull(rss!Description), "", rss!Description) & "'," _
                & "entry_date = '" & IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "company_code = '" & rss!COMPANY_CODE & "' " _
                & "WHERE company_code = '" & TDBCombo_company.Text & "' AND general_leave_number = '" & rss!general_leave_number & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_general_leave = "Total Data Cuti Bersama = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataLeavePeriode()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("t_leave_periode") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM t_leave_periode"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from t_leave_periode", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_leave_periode(rss!employee_code, rss!start_periode, rss!end_periode) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("employee_code").Value = rss!employee_code
                .Fields("start_working").Value = IIf(rss!start_working = "", "00:00:00", Format(rss!start_working, "yyyy-mm-dd HH:mm:ss"))
                .Fields("start_periode").Value = IIf(rss!start_periode = "", "00:00:00", Format(rss!start_periode, "yyyy-mm-dd HH:mm:ss"))
                .Fields("end_periode").Value = IIf(rss!end_periode = "", "00:00:00", Format(rss!end_periode, "yyyy-mm-dd HH:mm:ss"))
                .Fields("max_leave").Value = rss!max_leave
                .Fields("actual_leave").Value = rss!actual_leave
                .Fields("over_leave").Value = rss!over_leave
                .Fields("flag_close").Value = rss!flag_close
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE t_leave_periode SET employee_code = '" & rss!employee_code & "'," _
                & "start_working = '" & IIf(rss!start_working = "", "00:00:00", Format(rss!start_working, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "start_periode = '" & IIf(rss!start_periode = "", "00:00:00", Format(rss!start_periode, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "end_periode = '" & IIf(rss!end_periode = "", "00:00:00", Format(rss!end_periode, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "max_leave = '" & rss!max_leave & "'," _
                & "actual_leave = '" & rss!actual_leave & "'," _
                & "over_leave = '" & rss!over_leave & "'," _
                & "flag_close = '" & rss!flag_close & "' " _
                & "WHERE employee_code = '" & rss!employee_code & "' AND date(start_periode) = '" & Format(rss!start_periode, "yyyy-MM-dd") & "' AND " _
                & "date(end_periode) = '" & Format(rss!end_periode, "yyyy-MM-dd") & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    v_leave_periode = "Total Data Periode Cuti = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataSalaryStandard()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_salary_standard") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM m_salary_standard"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_salary_standard", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_salary_standard(rss!employee_code, rss!salary_date) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("main_salary").Value = rss!main_salary
                .Fields("staff_allowance").Value = rss!staff_allowance
                .Fields("functional_allowance").Value = rss!functional_allowance
                .Fields("phone_allowance").Value = rss!phone_allowance
                .Fields("transport_allowance").Value = rss!transport_allowance
                .Fields("other_allowance").Value = rss!other_allowance
                .Fields("presence_allowance").Value = rss!presence_allowance
                .Fields("meal_allowance").Value = rss!meal_allowance
                .Fields("special_allowance").Value = rss!special_allowance
                .Fields("employee_code").Value = rss!employee_code
                .Fields("driver_allowance").Value = rss!driver_allowance
                .Fields("acting_allowance").Value = rss!acting_allowance
                .Fields("skill_allowance").Value = rss!skill_allowance
                .Fields("entry_date").Value = IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss"))
                .Fields("user_entry").Value = rss!user_entry
                .Fields("pph21_type").Value = rss!pph21_type
                .Fields("ptkp_type").Value = rss!ptkp_type
                .Fields("jstk_type").Value = rss!jstk_type
                .Fields("flag_ot").Value = rss!flag_ot
                .Fields("salary_date").Value = rss!salary_date
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE m_salary_standard SET main_salary = '" & rss!main_salary & "'," _
                & "staff_allowance = '" & rss!staff_allowance & "'," _
                & "functional_allowance = '" & rss!functional_allowance & "'," _
                & "phone_allowance = '" & rss!phone_allowance & "'," _
                & "transport_allowance = '" & rss!transport_allowance & "'," _
                & "other_allowance = '" & rss!other_allowance & "'," _
                & "presence_allowance = '" & rss!presence_allowance & "'," _
                & "meal_allowance = '" & rss!meal_allowance & "'," _
                & "special_allowance = '" & rss!special_allowance & "'," _
                & "employee_code = '" & rss!employee_code & "'," _
                & "driver_allowance = '" & rss!driver_allowance & "'," _
                & "acting_allowance = '" & rss!acting_allowance & "'," _
                & "skill_allowance = '" & rss!skill_allowance & "'," _
                & "entry_date = '" & IIf(rss!entry_date = "", "00:00:00", Format(rss!entry_date, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "user_entry = '" & rss!user_entry & "'," _
                & "pph21_type = '" & rss!pph21_type & "'," _
                & "ptkp_type = '" & rss!ptkp_type & "'," _
                & "jstk_type = '" & rss!jstk_type & "'," _
                & "flag_ot = '" & rss!flag_ot & "'," _
                & "salary_date = '" & rss!salary_date & "' " _
                & "WHERE employee_code = '" & rss!employee_code & "' " _
                & "AND salary_date = '" & rss!salary_date & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    rs.Close
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    
    v_salary_standard = "Total Salary Standard = " & i & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataOtherIncome()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("t_employee_income") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM t_employee_income"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from t_employee_income", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_other_income(rss!noTrans, rss!employee_code) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("nomer").Value = rss!nomer
                .Fields("notrans").Value = rss!noTrans
                .Fields("tgltrans").Value = rss!tgltrans
                .Fields("employee_code").Value = rss!employee_code
                .Fields("jmlPotong").Value = rss!jmlpotong
                .Fields("remark").Value = rss!remark
                .Fields("userinput").Value = rss!userInput
                .Fields("tglinput").Value = IIf(rss!tglinput = "", "00:00:00", Format(rss!tglinput, "yyyy-mm-dd HH:mm:ss"))
                .Fields("useredit").Value = rss!userEdit
                .Fields("tgledit").Value = IIf(rss!tglEdit = "", "00:00:00", Format(rss!tglEdit, "yyyy-mm-dd HH:mm:ss"))
                .Fields("flag_other_income").Value = rss!flag_other_income
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE t_employee_income SET nomer = '" & rss!nomer & "'," _
                & "notrans = '" & rss!noTrans & "'," _
                & "tgltrans = '" & rss!tgltrans & "'," _
                & "employee_code = '" & rss!employee_code & "'," _
                & "jmlPotong = '" & rss!jmlpotong & "'," _
                & "remark = '" & rss!remark & "'," _
                & "userinput = '" & rss!userInput & "'," _
                & "tglinput = '" & IIf(rss!tglinput = "", "00:00:00", Format(rss!tglinput, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "useredit = '" & rss!userEdit & "'," _
                & "tgledit = '" & IIf(rss!tglEdit = "", "00:00:00", Format(rss!tglEdit, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "flag_other_income = '" & rss!flag_other_income & "' " _
                & "WHERE notrans = '" & rss!noTrans & "' " _
                & "AND employee_code = '" & rss!employee_code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    rs.Close
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    
    v_other_income = "Total Other Income = " & i & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataOtherExpense()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("t_employee_expense") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM t_employee_expense"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from t_employee_expense", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_other_expense(rss!noTrans, rss!employee_code) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("nomer").Value = rss!nomer
                .Fields("notrans").Value = rss!noTrans
                .Fields("tgltrans").Value = rss!tgltrans
                .Fields("employee_code").Value = rss!employee_code
                .Fields("jmlPotong").Value = rss!jmlpotong
                .Fields("remark").Value = rss!remark
                .Fields("userinput").Value = rss!userInput
                .Fields("tglinput").Value = IIf(rss!tglinput = "", "00:00:00", Format(rss!tglinput, "yyyy-mm-dd HH:mm:ss"))
                .Fields("useredit").Value = rss!userEdit
                .Fields("tgledit").Value = IIf(rss!tglEdit = "", "00:00:00", Format(rss!tglEdit, "yyyy-mm-dd HH:mm:ss"))
                .Fields("flag_other_expense").Value = rss!flag_other_expense
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE t_employee_expense SET nomer = '" & rss!nomer & "'," _
                & "notrans = '" & rss!noTrans & "'," _
                & "tgltrans = '" & rss!tgltrans & "'," _
                & "employee_code = '" & rss!employee_code & "'," _
                & "jmlPotong = '" & rss!jmlpotong & "'," _
                & "remark = '" & rss!remark & "'," _
                & "userinput = '" & rss!userInput & "'," _
                & "tglinput = '" & IIf(rss!tglinput = "", "00:00:00", Format(rss!tglinput, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "useredit = '" & rss!userEdit & "'," _
                & "tgledit = '" & IIf(rss!tglEdit = "", "00:00:00", Format(rss!tglEdit, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "flag_other_expense = '" & rss!flag_other_expense & "' " _
                & "WHERE notrans = '" & rss!noTrans & "' " _
                & "AND employee_code = '" & rss!employee_code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    rs.Close
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    
    v_other_expense = "Total Other Expense = " & i & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataLoan()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("tm_loan") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM tm_loan"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tm_loan", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_tm_loan(rss!employee_code, rss!loan_date) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("date").Value = rss!Date
                .Fields("employee_code").Value = rss!employee_code
                .Fields("loan_value").Value = rss!loan_value
                .Fields("loan_interest").Value = rss!loan_interest
                .Fields("loan_total").Value = rss!loan_total
                .Fields("installment_time").Value = rss!installment_time
                .Fields("installment_value").Value = rss!installment_value
                .Fields("installment_start").Value = IIf(rss!installment_start = "", "00:00:00", Format(rss!installment_start, "yyyy-mm-dd HH:mm:ss"))
                .Fields("installment_end").Value = IIf(rss!installment_end = "", "00:00:00", Format(rss!installment_end, "yyyy-mm-dd HH:mm:ss"))
                .Fields("description").Value = rss!Description
                .Fields("company_code").Value = rss!COMPANY_CODE
                .Fields("employee_name").Value = rss!EMPLOYEE_NAME
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE t_loan SET employee_code = '" & rss!employee_code & "'," _
                & "loan_date = '" & rss!loan_date & "'," _
                & "loan_value = '" & rss!loan_value & "'," _
                & "loan_interest = '" & rss!loan_interest & "'," _
                & "loan_total = '" & rss!loan_total & "'," _
                & "installment_time = '" & rss!installment_time & "'," _
                & "installment_value = '" & rss!installment_value & "'," _
                & "installment_start = '" & IIf(rss!installment_start = "", "00:00:00", Format(rss!installment_start, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "installment_end = '" & IIf(rss!installment_end = "", "00:00:00", Format(rss!installment_end, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "description = '" & rss!Description & "',company_code = '" & rss!COMPANY_CODE & "', " _
                & "employee_name = '" & rss!EMPLOYEE_NAME & "' " _
                & "WHERE employee_code = '" & rss!employee_code & "' " _
                & "AND date(loan_date) = '" & Format(rss!loan_date, "yyyy-MM-dd") & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    rs.Close
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    
    v_loan = "Total Loan = " & i & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataLoanDetail()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("td_loan") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM td_loan"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from td_loan", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_td_loan(rss!employee_code, rss!loan_date, rss!sequence_number) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("date").Value = rss!Date
                .Fields("employee_date").Value = rss!employee_date
                .Fields("salary_item_code").Value = rss!salary_item_code
                .Fields("sequence_number").Value = rss!sequence_number
                .Fields("employee_name").Value = rss!EMPLOYEE_NAME
                .Fields("salary_item_name").Value = rss!salary_item_name
                .Fields("installment_month").Value = IIf(rss!installment_month = "", "00:00:00", Format(rss!installment_month, "yyyy-mm-dd HH:mm:ss"))
                .Fields("installment_amount").Value = rss!installment_amount
                .Fields("installment_pay").Value = rss!installment_pay
                .Fields("installment_pay_date").Value = IIf(rss!installment_pay_date = "", "00:00:00", Format(rss!installment_pay_date, "yyyy-mm-dd HH:mm:ss"))
                .Fields("flag_paid").Value = rss!flag_paid
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE td_loan SET date = '" & rss!Date & "'," _
                & "employee_date = '" & rss!employee_date & "'," _
                & "salary_item_code = '" & rss!salary_item_code & "'," _
                & "sequence_number = '" & rss!sequence_number & "'," _
                & "employee_name = '" & rss!EMPLOYEE_NAME & "'," _
                & "salary_item_name = '" & rss!salary_item_name & "'," _
                & "installment_month = '" & IIf(rss!installment_month = "", "00:00:00", Format(rss!installment_month, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "installment_amount = '" & rss!installment_amount & "'," _
                & "installment_pay = '" & rss!installment_pay & "'," _
                & "installment_pay_date = '" & IIf(rss!installment_pay_date = "", "00:00:00", Format(rss!installment_pay_date, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "flag_paid = '" & rss!flag_paid & "' " _
                & "WHERE employee_code = '" & rss!employee_code & "' " _
                & "AND date(date) = '" & Format(rss!Date, "yyyy-MM-dd") & "' " _
                & "AND sequence_number = '" & rss!sequence_number & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    rs.Close
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    
    v_loan_detail = "Total Loan Detail = " & i & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataUser()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_user") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM m_user"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
'        If rss.RecordCount > 0 Then
'            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
'                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
'                Exit Sub
'            End If
'        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_user", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_m_user(rss!user_code) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("user_code").Value = rss!user_code
                .Fields("company_code").Value = rss!COMPANY_CODE
                .Fields("user_name").Value = rss!USER_NAME
                .Fields("user_pass").Value = rss!USER_PASS
                .Fields("user_pass_key").Value = rss!user_pass_key
                .Fields("user_level").Value = rss!user_level
                .Fields("flag_company_access").Value = rss!flag_company_access
                .Fields("employee_code").Value = rss!employee_code
                .Fields("flag_user").Value = rss!flag_user
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE m_user SET user_code = '" & rss!user_code & "'," _
                & "company_code = '" & rss!COMPANY_CODE & "'," _
                & "user_name = '" & rss!USER_NAME & "'," _
                & "user_pass = '" & rss!USER_PASS & "'," _
                & "user_pass_key = '" & rss!user_pass_key & "'," _
                & "user_level = '" & rss!user_level & "'," _
                & "flag_company_access = '" & rss!flag_company_access & "'," _
                & "employee_code = '" & rss!employee_code & "'," _
                & "flag_user = '" & rss!flag_user & "' " _
                & "WHERE user_code = '" & rss!user_code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    rs.Close
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    
'    v_other_expense = "Total Other Expense = " & i & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataUserDetail()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("t_user") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM t_user"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
'        If rss.RecordCount > 0 Then
'            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
'                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
'                Exit Sub
'            End If
'        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from t_user", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_t_user(rss!level_code, rss!sub_menu_code) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("level_code").Value = rss!level_code
                .Fields("sub_menu_code").Value = rss!sub_menu_code
                .Fields("sub_menu_name").Value = rss!sub_menu_name
                .Fields("menu_code").Value = rss!menu_code
                .Fields("menu_name").Value = rss!menu_name
                .Fields("form_name").Value = rss!form_name
                .Fields("form_title").Value = rss!form_title
                .Fields("allow_read").Value = rss!allow_read
                .Fields("allow_add").Value = rss!allow_add
                .Fields("allow_edit").Value = rss!allow_edit
                .Fields("allow_delete").Value = rss!allow_delete
                .Fields("allow_post").Value = rss!allow_post
                .Fields("allow_print").Value = rss!allow_print
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE t_user SET level_code = '" & rss!level_code & "'," _
                & "sub_menu_code = '" & rss!sub_menu_code & "'," _
                & "sub_menu_name = '" & Replace(rss!sub_menu_name, "'", "''") & "'," _
                & "menu_code = '" & rss!menu_code & "'," _
                & "menu_name = '" & Replace(rss!menu_name, "'", "''") & "'," _
                & "form_name = '" & rss!form_name & "'," _
                & "form_title = '" & rss!form_title & "'," _
                & "allow_read = '" & rss!allow_read & "'," _
                & "allow_add = '" & rss!allow_add & "'," _
                & "allow_edit = '" & rss!allow_edit & "'," _
                & "allow_delete = '" & rss!allow_delete & "'," _
                & "allow_post = '" & rss!allow_post & "'," _
                & "allow_print = '" & rss!allow_print & "' " _
                & "WHERE level_code = '" & rss!level_code & "' " _
                    & "AND sub_menu_code = '" & rss!sub_menu_code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    rs.Close
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    
'    v_other_expense = "Total Other Expense = " & i & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataUserAccess()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("t_user_access_level") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM t_user_access_level"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
'        If rss.RecordCount > 0 Then
'            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
'                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
'                Exit Sub
'            End If
'        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from t_user_access_level", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_t_user_access(rss!level_code, rss!access_level_code) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("level_code").Value = rss!level_code
                .Fields("access_level_code").Value = rss!access_level_code
                .Fields("level_name").Value = rss!level_name
                .Fields("allow_access").Value = rss!allow_access
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE t_user_access_level SET level_code = '" & rss!level_code & "'," _
                & "access_level_code = '" & rss!access_level_code & "'," _
                & "level_name = '" & rss!level_name & "'," _
                & "allow_access = '" & rss!allow_access & "' " _
                & "WHERE level_code = '" & rss!level_code & "' " _
                    & "AND access_level_code = '" & rss!access_level_code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    rs.Close
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    
    Set rss = Nothing
    
'    v_other_expense = "Total Other Expense = " & i & " !"
'    MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    Exit Sub
End Sub

Private Sub importDataShift()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_shift") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM m_shift"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_shift", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_shift(rss!shift_code, rss!COMPANY_CODE) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("shift_code").Value = rss!shift_code
                .Fields("shift_name").Value = rss!shift_name
                .Fields("start_time").Value = rss!start_time
                .Fields("end_time").Value = rss!end_time
                .Fields("flag_day_over").Value = rss!flag_day_over
                .Fields("flag_tolerance").Value = rss!flag_tolerance
                .Fields("start_time_tolerance").Value = IIf(rss!start_time_tolerance = "", "00:00:00", Format(rss!start_time_tolerance, "yyyy-mm-dd HH:mm:ss"))
                .Fields("end_time_tolerance").Value = IIf(rss!end_time_tolerance = "", "00:00:00", Format(rss!end_time_tolerance, "yyyy-mm-dd HH:mm:ss"))
                .Fields("flag_shift").Value = rss!flag_shift
                .Fields("min_break_in").Value = rss!min_break_in
                .Fields("max_break_out").Value = rss!max_break_out
                .Fields("break_interval_minute").Value = rss!break_interval_minute
                .Fields("flag_moving").Value = rss!flag_moving
                .Fields("moving_number").Value = rss!moving_number
                .Fields("company_code").Value = rss!COMPANY_CODE
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE m_shift SET shift_code = '" & rss!shift_code & "'," _
                & "shift_name = '" & rss!shift_name & "'," _
                & "start_time = '" & rss!start_time & "'," _
                & "end_time = '" & rss!end_time & "'," _
                & "flag_day_over = '" & rss!flag_day_over & "'," _
                & "flag_tolerance = '" & rss!flag_tolerance & "'," _
                & "start_time_tolerance = '" & IIf(rss!start_time_tolerance = "", "00:00:00", Format(rss!start_time_tolerance, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "end_time_tolerance = '" & IIf(rss!end_time_tolerance = "", "00:00:00", Format(rss!end_time_tolerance, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "flag_shift = '" & rss!flag_shift & "'," _
                & "min_break_in = '" & rss!min_break_in & "'," _
                & "max_break_out = '" & rss!max_break_out & "'," _
                & "break_interval_minute = '" & rss!break_interval_minute & "'," _
                & "flag_moving = '" & rss!flag_moving & "'," _
                & "moving_number = '" & rss!moving_number & "'," _
                & "company_code = '" & rss!COMPANY_CODE & "' " _
                & "WHERE shift_code = '" & rss!shift_code & "' AND company_code = '" & rss!COMPANY_CODE & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    rs.Close
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    v_shift = "Total Data Shift = " & ProgressBar1.Value & " !"
    Exit Sub
End Sub

Private Sub importDataSalary()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("h_salary") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM h_salary"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
        If rss.RecordCount > 0 Then
            If rss!COMPANY_CODE <> TDBCombo_company.Text Then
                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
                Exit Sub
            End If
        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from h_salary", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_new_salary(rss!month, rss!employee_code, rss!salary_code) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("month").Value = rss!month
                .Fields("employee_code").Value = rss!employee_code
                .Fields("salary_code").Value = rss!salary_code
                .Fields("salary_name").Value = rss!salary_name
                .Fields("date_from").Value = IIf(rss!date_from = "", "00:00:00", Format(rss!date_from, "yyyy-mm-dd HH:mm:ss"))
                .Fields("date_to").Value = IIf(rss!date_to = "", "00:00:00", Format(rss!date_to, "yyyy-mm-dd HH:mm:ss"))
                .Fields("flag_main_salary").Value = rss!flag_main_salary
                .Fields("flag_sign").Value = rss!flag_sign
                .Fields("flag_detail").Value = rss!flag_detail
                .Fields("flag_use_formula").Value = rss!flag_use_formula
                .Fields("formula_salary_code").Value = rss!formula_salary_code
                .Fields("flag_ptkp").Value = rss!flag_ptkp
                .Fields("ptkp_salary_code").Value = rss!ptkp_salary_code
                .Fields("flag_pkp").Value = rss!flag_pkp
                .Fields("flag_pph21").Value = rss!flag_pph21
                .Fields("pph21_number").Value = rss!pph21_number
                .Fields("flag_tax").Value = rss!flag_tax
                .Fields("tax_salary_code").Value = rss!tax_salary_code
                .Fields("flag_type").Value = rss!flag_type
                .Fields("flag_visible").Value = rss!flag_visible
                .Fields("salary_value").Value = rss!salary_value
                .Fields("description").Value = rss!Description
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE h_salary SET month = '" & rss!month & "'," _
                & "employee_code = '" & rss!employee_code & "'," _
                & "salary_code = '" & rss!salary_code & "'," _
                & "salary_name = '" & rss!salary_name & "'," _
                & "date_from = '" & IIf(rss!date_from = "", "00:00:00", Format(rss!date_from, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "date_to = '" & IIf(rss!date_to = "", "00:00:00", Format(rss!date_to, "yyyy-mm-dd HH:mm:ss")) & "'," _
                & "flag_main_salary = '" & rss!flag_main_salary & "'," _
                & "flag_sign = '" & rss!flag_sign & "'," _
                & "flag_detail = '" & rss!flag_detail & "'," _
                & "flag_use_formula = '" & rss!flag_use_formula & "'," _
                & "formula_salary_code = '" & rss!formula_salary_code & "'," _
                & "flag_ptkp = '" & rss!flag_ptkp & "'," _
                & "ptkp_salary_code = '" & rss!ptkp_salary_code & "'," _
                & "flag_pkp = '" & rss!flag_pkp & "'," _
                & "flag_pph21 = '" & rss!flag_pph21 & "'," _
                & "pph21_number = '" & rss!pph21_number & "'," _
                & "flag_tax = '" & rss!flag_tax & "'," _
                & "tax_salary_code = '" & rss!tax_salary_code & "'," _
                & "flag_type = '" & rss!flag_type & "'," _
                & "flag_visible = '" & rss!flag_visible & "'," _
                & "salary_value = '" & rss!salary_value & "'," _
                & "description = '" & rss!Description & "' " _
                & "WHERE month = '" & rss!month & "' AND employee_code = '" & rss!employee_code & "' AND " _
                & "salary_code = '" & rss!salary_code & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    rs.Close
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    Exit Sub
End Sub
Private Sub importDataHDSalary()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("h_d_salary") Then
        MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
        Exit Sub
    Else
        strsql = "SELECT * FROM h_d_salary"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
        
'        If rss.RecordCount > 0 Then
'            If rss!company_code <> TDBCombo_company.Text Then
'                MsgBox "File Tidak Valid Untuk Perusahaan Ini, Cek Ulang File...", vbExclamation, headerMSG
'                Exit Sub
'            End If
'        End If
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from h_d_salary", CnG, adOpenKeyset, adLockOptimistic
    
'    Label2.Visible = True
    ProgressBar1.Visible = True
    
    If rss.RecordCount > 0 Then
        ProgressBar1.Max = rss.RecordCount
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0
    End If
        
    For aa = 0 To rss.RecordCount - 1
    
        DoEvents
        ProgressBar1.Value = aa
'        Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
    
        If Not check_exist_h_d_salary(rss!month, rss!COMPANY_CODE) Then
            i = i + 1
                    
            With rs
                .AddNew
                '-----------------------
                .Fields("month").Value = rss!month
                .Fields("periode_from").Value = rss!periode_from
                .Fields("periode_to").Value = rss!periode_to
                .Fields("company_code").Value = rss!COMPANY_CODE
                .Fields("company_name").Value = rss!company_name
                '-----------------------
                .Update
            End With
        Else
            strsql = "UPDATE h_d_salary SET month = '" & rss!month & "'," _
                & "periode_from = '" & rss!periode_from & "'," _
                & "periode_to = '" & rss!periode_to & "'," _
                & "company_code = '" & rss!COMPANY_CODE & "'," _
                & "company_name = '" & rss!company_name & "' " _
                & "WHERE date(month) = '" & Format(rss!month, "yyyy-MM-dd") & "' AND " _
                & "company_code = '" & rss!COMPANY_CODE & "'"
            CnG.Execute strsql
        End If
        rss.MoveNext
    Next
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    
'    Label2.Visible = False
    ProgressBar1.Visible = False
    'MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
'    v_department = "Total Data Departement / Area = " & i & " !"
    Exit Sub
End Sub

'**********Prevent window close if table content being load
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Frame1.Enabled Then
        Cancel = 0
    Else
        Cancel = 1
    End If
End Sub

Public Function SelectTable(tableName As String) As Boolean
Dim rss As New cRecordset
Dim strsql As String
On Error GoTo ErrHandler
    strsql = "SELECT 1 FROM " & tableName

    rss.OpenRecordset strsql, conLite, True
    
    SelectTable = True
Exit Function

ErrHandler:
'MsgBox "Tidak Bisa Menemukan Data Pada File Yang Dipilih.. Cek Ulang File..."
SelectTable = False

End Function

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
    
    Frame1.Enabled = True
    Frame2.Enabled = True
    
    Call load_data_department
End If

'If TDBCombo_company.Text = "GPN" Then
'    DTPicker_periode_from = Format(Now, "yyyy-MM-") & "01"
'    DTPicker_periode_to = Format(Now, "yyyy-MM-") & getEndDay(month(Now), year(Now))
'Else
'    DTPicker_periode_from = Format(DateAdd("M", -1, Now), "yyyy-MM-") & "26"
'    DTPicker_periode_to = Format(Now, "yyyy-MM-") & "25"
'End If

End Sub

Private Sub TDBCombo_department_itemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
    txt_department_name = TDBCombo_department.Columns("department_name").Value
    
    If optType(0) And OptData(0) Then
        LynxGrid1.ClearAll
        Call createGridEmployee
        Call fillTableEmployee
    End If
End If
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub load_data_department()
Adodc_department.RecordSource = "select * from m_department where company_code = '" & TDBCombo_company.Text & "' order by department_code"
Adodc_department.Refresh

TDBCombo_department.RowSource = Adodc_department

End Sub

Private Sub Timer1_Timer()
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub


