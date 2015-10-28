VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_trans_tranfer_data 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EXPORT - IMPORT DATA"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11505
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   465
      Left            =   9120
      TabIndex        =   24
      Top             =   6990
      Width           =   1935
   End
   Begin VB.Frame frmLoading 
      Height          =   555
      Left            =   3930
      TabIndex        =   21
      Top             =   3150
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
         TabIndex        =   22
         Top             =   180
         Width           =   3465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   1260
      TabIndex        =   16
      Top             =   600
      Width           =   4365
      Begin VB.OptionButton OptData 
         Caption         =   "Data Attendance"
         Height          =   315
         Index           =   1
         Left            =   2070
         TabIndex        =   20
         Top             =   210
         Width           =   1665
      End
      Begin VB.OptionButton OptData 
         Caption         =   "Master Employee"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   1260
      TabIndex        =   15
      Top             =   60
      Width           =   4365
      Begin VB.OptionButton OptType 
         Caption         =   "Import"
         Height          =   285
         Index           =   1
         Left            =   2070
         TabIndex        =   18
         Top             =   180
         Width           =   2175
      End
      Begin VB.OptionButton OptType 
         Caption         =   "Export"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
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
      Left            =   9570
      TabIndex        =   6
      Top             =   1500
      Width           =   225
   End
   Begin prj_fej_jkt.LynxGrid LynxGrid1 
      Height          =   3645
      Left            =   420
      TabIndex        =   5
      Top             =   1860
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   6429
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
   Begin prj_fej_jkt.vbButton btnBrowse 
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Top             =   6060
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
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      TabIndex        =   2
      Top             =   6060
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4050
      Top             =   6870
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
      Top             =   6450
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
      Top             =   6990
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker_periode_from 
      Height          =   300
      Left            =   1260
      TabIndex        =   10
      Top             =   1350
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   126877699
      CurrentDate     =   39278
      MaxDate         =   402133
      MinDate         =   36526
   End
   Begin MSComCtl2.DTPicker DTPicker_periode_to 
      Height          =   300
      Left            =   3930
      TabIndex        =   11
      Top             =   1350
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   126877699
      CurrentDate     =   39278
      MaxDate         =   402133
      MinDate         =   36526
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
      TabIndex        =   23
      Top             =   6780
      Width           =   3465
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA"
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Top             =   870
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
      TabIndex        =   13
      Top             =   1380
      Width           =   615
   End
   Begin VB.Label lblPeriode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE"
      Height          =   195
      Left            =   450
      TabIndex        =   12
      Top             =   1410
      Width           =   720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE"
      Height          =   345
      Left            =   480
      TabIndex        =   9
      Top             =   300
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
      TabIndex        =   8
      Top             =   450
      Width           =   5295
   End
   Begin VB.Label lblAll 
      BackStyle       =   0  'Transparent
      Caption         =   "Select All"
      Height          =   255
      Left            =   9990
      TabIndex        =   7
      Top             =   1470
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
      Top             =   5640
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

Private Sub cbo_periode_to_Change()
'Call periode_date_event
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

If Len(txtFileName.Text) = 0 Then
    MsgBox "Invalid File Name....!!!", vbCritical
    Exit Sub
End If

If OptType(0) Then ' Export Transfer Type

    If Right(txtFileName.Text, 4) = ".ssd" Then
        nmFile = txtFileName.Text
    Else
        nmFile = txtFileName.Text & ".ssd"
    End If

    If checkFile(nmFile) Then
        i = MsgBox("File has been existed ! " & _
            "Overwrite File ?", vbOKCancel, headerMSG)
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
            MsgBox "No Data Selected !", vbInformation, headerMSG
            Exit Sub
        End If
        cmdRun.Enabled = False
        Call exportDataEmployee
        cmdRun.Enabled = True
    Else ' Attendance Data
        cmdRun.Enabled = False
        Call exportDataAttendance
        cmdRun.Enabled = True
    End If
Else ' Import Transfer Type
    
    If Not Right(txtFileName.Text, 4) = ".ssd" Then
        MsgBox "Invalid file name ! Filename format must be in <filename>.ssd !"
        Exit Sub
    End If
    
    If Not checkFile(txtFileName.Text) Then
        MsgBox "File not found !"
        Exit Sub
    End If
    
    cmdRun.Enabled = False
    
    If OptData(0) Then ' Master Employee Data
        Call importDataEmployee
    Else ' Attendance Data
        Call importDataAttendance
    End If
    
    cmdRun.Enabled = True
End If

End Sub

Private Sub Command1_Click()
Unload Me
Exit Sub
End Sub

Private Sub Form_Load()
frmLoading.Visible = False
Label2.Visible = False

DTPicker_periode_from = Now
DTPicker_periode_to = Now

'Call createGridEmployee
'Call fillTableEmployee

OptType(1) = True
OptData(0) = True

End Sub


Private Sub OptType_Click(Index As Integer)
   
    Call Option_Click
    
End Sub

Private Sub OptData_Click(Index As Integer)
    
    Call Option_Click
    
End Sub

Private Sub Option_Click()
    
    If OptData(0) Then
        lblTitleGrid.Caption = " MASTER EMPLOYEE LIST"
        Call mode_periode(False)
        If OptType(0) Then
            lblFileTo.Caption = " Export File To : "
            LynxGrid1.Enabled = True
            Call createGridEmployee
                Frame1.Enabled = False
                Frame2.Enabled = False

            If fillTableEmployee() Then
                Frame1.Enabled = True
                Frame2.Enabled = True
            End If
        Else
            lblFileTo.Caption = " Import File From : "
            If LynxGrid1.Rows > 0 Then
                LynxGrid1.ClearAll
            End If
            LynxGrid1.Enabled = False
            lblAll.Visible = False
            ChkAll.Visible = False
        End If
    Else
        lblTitleGrid.Caption = " ATTENDANCE DATA LIST "
        If LynxGrid1.Rows > 0 Then
            LynxGrid1.ClearAll
        End If
        If OptType(0) Then
            Call mode_periode(True)
            lblFileTo.Caption = " Export File To : "
            LynxGrid1.Enabled = True
        Else
            Call mode_periode(False)
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
If OptType(0) Then
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
      .AddColumn "Div. Code", , , , , , , , , False
      .AddColumn "Div. Name", 2500, , , , , , , , , True
      .AddColumn "Title Code", , , , , , , , , False
      .AddColumn "Position", 2500, , , , , , , , , True
      .AddColumn "Check", 800, lgAlignCenterCenter, lgBoolean, , , , , , True, False
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
    
    strsql = "SELECT employee_code,employee_name,division_code,division_name,title_code, title_name " & _
        "FROM m_employee WHERE flag_active = 1"
    
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!employee_code & vbTab & rs!employee_name & _
                vbTab & rs!division_code & vbTab & rs!division_name & _
                vbTab & rs!title_code & vbTab & rs!title_name & _
                vbTab & 0
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
    lblPeriode.Visible = str
    DTPicker_periode_from.Visible = str
    DTPicker_periode_to.Visible = str
    lblRange.Visible = str
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

Private Sub exportDataEmployee()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

    conLite.Execute "Create Table If Not Exists m_employee (" & _
                    "employee_code  TEXT(20) NOT NULL, employee_name  TEXT(50) DEFAULT NULL, employee_nick_name  TEXT(50) DEFAULT NULL, " & _
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
                    "flag_shiftable  INTEGER DEFAULT NULL,flag_active  INTEGER DEFAULT NULL,description  TEXT(100) DEFAULT NULL, " & _
                    "end_working  TEXT(100) DEFAULT NULL,reason  TEXT(100) DEFAULT NULL,flag_pic  INTEGER DEFAULT NULL, " & _
                    "pic  BLOB,access_code  TEXT(10) DEFAULT NULL,managerial_level  INTEGER DEFAULT NULL, " & _
                    "fathers_name  TEXT(50) DEFAULT NULL,mothers_name  TEXT(50) DEFAULT NULL,child_number_from  INTEGER DEFAULT NULL, " & _
                    "child_number  INTEGER DEFAULT NULL," & _
                    "start_mc  TEXT(30) DEFAULT NULL,end_mc  TEXT(30) DEFAULT NULL,branches  TEXT(50) DEFAULT NULL," & _
                    "PRIMARY KEY (employee_code ASC)); "
    
    strsql = "INSERT INTO m_employee " & _
                "(employee_code,employee_name,employee_nick_name,division_code,division_name," & _
                "department_code,department_name,company_code,company_name,date_of_birth,place_of_birth," & _
                "sex,religion,marital_status,number_of_children,address,email,npwp,phone_number,bank_account," & _
                "last_education_code,last_education_code_other,last_education_name,last_education_pass,last_employment_name," & _
                "last_employment_date,last_employment_title,start_working,date_of_appointment,title_code,title_name,flag_shiftable," & _
                "flag_active,description,end_working,reason,flag_pic,pic,access_code,managerial_level,fathers_name,mothers_name," & _
                "child_number_from,child_number,start_mc,end_mc,branches) " & _
                "Values " & _
                "(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    
    Set cmd = conLite.CreateCommand(strsql)
    
    bb = 0 'buat nyimpan jumlah data yang di select
    
    For aa = 0 To LynxGrid1.Rows - 1
        If LynxGrid1.CellValue(aa, 6) = True Then
            bb = bb + 1
        End If
    Next
    
'    If bb = 0 Then
'        MsgBox "No Data Selected !", vbInformation, headerMSG
'        Exit Sub
'    End If
    
    ProgressBar1.Max = bb
    ProgressBar1.Value = 0
    
    ProgressBar1.Visible = True
    Label2.Visible = True
    
    If LynxGrid1.Rows > 0 Then
        conLite.BeginTrans
        conLite.Execute ("DELETE FROM m_employee")
        For aa = 0 To LynxGrid1.Rows - 1
                        
            If LynxGrid1.CellValue(aa, 6) = True Then

                strsql = "SELECT * FROM m_employee WHERE employee_code = '" & LynxGrid1.CellText(aa, 0) & "' "
                Set rsemp = New ADODB.Recordset
                
                'If rsEmp.State = 1 Then rsEmp.Close
                rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                DoEvents
                ProgressBar1.Value = ProgressBar1.Value + 1
                Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                
                'Set rsLite = conLite.OpenRecordset(strSQl, True)
                cmd.SetText 1, rsemp!employee_code
                cmd.SetText 2, rsemp!employee_name
                cmd.SetText 3, IIf(IsNull(rsemp!employee_nick_name), "", rsemp!employee_nick_name)
                cmd.SetText 4, rsemp!division_code
                cmd.SetText 5, rsemp!division_name
                cmd.SetText 6, rsemp!department_code
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
                cmd.SetText 30, rsemp!title_code
                cmd.SetText 31, rsemp!title_name
                cmd.SetInt32 32, IIf(IsNull(rsemp!flag_shiftable), 0, rsemp!flag_shiftable)
                cmd.SetInt32 33, IIf(IsNull(rsemp!flag_active), 0, rsemp!flag_active)
                cmd.SetText 34, IIf(IsNull(rsemp!Description), "", rsemp!Description)
                cmd.SetDate 35, IIf(IsNull(rsemp!end_working) = True, 0, rsemp!end_working)
                cmd.SetText 36, IIf(IsNull(rsemp!reason) = True, "", rsemp!reason)
                cmd.SetInt32 37, IIf(IsNull(rsemp!flag_pic) = True, 0, rsemp!flag_pic)
                cmd.SetText 38, "" ' rsEmp!pic
                cmd.SetText 39, IIf(IsNull(rsemp!access_code) = True, "", rsemp!access_code)
                cmd.SetInt32 40, IIf(IsNull(rsemp!managerial_level), 0, rsemp!managerial_level)
                cmd.SetText 41, IIf(IsNull(rsemp!fathers_name), "", rsemp!fathers_name)
                cmd.SetText 42, IIf(IsNull(rsemp!mothers_name), "", rsemp!mothers_name)
                cmd.SetInt32 43, IIf(IsNull(rsemp!child_number_from), 0, rsemp!child_number_from)
                cmd.SetInt32 44, IIf(IsNull(rsemp!child_number), 0, rsemp!child_number)
                cmd.SetDate 45, IIf(IsNull(rsemp!start_mc), 0, rsemp!start_mc)
                cmd.SetDate 46, IIf(IsNull(rsemp!end_mc), 0, rsemp!end_mc)
                cmd.SetText 47, IIf(IsNull(rsemp!branches), "", rsemp!branches)
                cmd.Execute
            End If
        Next
        conLite.CommitTrans
    End If
    
    MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    ProgressBar1.Visible = False
    Label2.Visible = False
    
End Sub

Private Sub exportDataAttendance()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

conLite.Execute "CREATE TABLE IF NOT EXISTS h_attendance (" & _
                        "employee_code TEXT(20) NOT NULL, att_date TEXT(30) NOT NULL, ip_address TEXT(50) DEFAULT NULL," & _
                        "enrollnumber INTEGER DEFAULT NULL, shift_number INTEGER DEFAULT NULL,shift_code TEXT(20) DEFAULT NULL," & _
                        "start_time TEXT(30) DEFAULT NULL,end_time TEXT(30) DEFAULT NULL,time_in TEXT(30) DEFAULT NULL," & _
                        "time_out TEXT(30) DEFAULT NULL,time_in_diff TEXT(30) DEFAULT NULL,time_out_diff TEXT(30) DEFAULT NULL," & _
                        "flag_io INTEGER DEFAULT NULL,flag_present INTEGER DEFAULT NULL,flag_duty INTEGER DEFAULT NULL," & _
                        "flag_late INTEGER DEFAULT NULL,flag_early INTEGER DEFAULT NULL,absent_status TEXT(5) DEFAULT NULL," & _
                        "time_in_break TEXT(30) DEFAULT NULL,time_in_break_diff TEXT(30) DEFAULT NULL,flag_in_break_early INTEGER DEFAULT NULL," & _
                        "time_out_break TEXT(30) DEFAULT NULL,time_out_break_diff TEXT(30) DEFAULT NULL,flag_out_break_late INTEGER DEFAULT NULL," & _
                        "break_interval TEXT(30) DEFAULT NULL,break_ot datetime DEFAULT NULL,description TEXT(50) DEFAULT NULL," & _
                        "entry_date TEXT(30) DEFAULT NULL,edit_date TEXT(30) DEFAULT NULL,userinput TEXT(10) DEFAULT NULL," & _
                        "useredit TEXT(10) DEFAULT NULL,PRIMARY KEY(employee_code, att_date)); "
                        
        strsql = "INSERT INTO h_attendance " & _
                    "(employee_code,att_date,ip_address,enrollnumber,shift_number," & _
                    "shift_code,start_time,end_time,time_in,time_out,time_in_diff," & _
                    "time_out_diff,flag_io,flag_present,flag_duty,flag_late,flag_early,absent_status," & _
                    "time_in_break,time_in_break_diff,flag_in_break_early,time_out_break,time_out_break_diff," & _
                    "flag_out_break_late,break_interval,break_ot,description,entry_date,edit_date,userinput," & _
                    "useredit) " & _
                    "Values " & _
                    "(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                                       
        Set cmd = conLite.CreateCommand(strsql)
        
        strsql = "SELECT * FROM h_attendance WHERE DATE(att_date) >= '" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "' " & _
                 "AND Date(att_date) <= '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "' "
                 
        Set rsAtt = New ADODB.Recordset
    
        rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsAtt.RecordCount > 0 Then
            ProgressBar1.Max = rsAtt.RecordCount
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Max = 1
            ProgressBar1.Value = 0
        End If
        
        ProgressBar1.Visible = True
        Label2.Visible = True
        
        'Set rs = Nothing
        
        'If LynxGrid1.Rows > 0 Then
            conLite.BeginTrans
            conLite.Execute ("DELETE FROM h_attendance")
            
            If Not rsAtt.EOF Then
                rsAtt.MoveFirst
                While Not rsAtt.EOF
                            
                'If LynxGrid1.CellValue(aa, 6) = True Then
                
        '            strsql = "SELECT * FROM h_attendance WHERE LEFT(att_date,10) >= '" & Format(DTPicker_periode_from.Value, "yyyy-mm-dd") & "' " & _
        '                    "AND LEFT(att_date,10) <= '" & Format(DTPicker_periode_to.Value, "yyyy-mm-dd") & "' " & _
        '                    "AND employee_code = '" & LynxGrid1.CellText(aa, 0) & "' AND att_date = '" & Format(LynxGrid1.CellText(aa, 1), "yyyy-mm-dd hh:mm:ss") & "' "
'
'                    Set rsAtt = New ADODB.Recordset
'                    rsAtt.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label2.Caption = "Total Data Transfered = " & ProgressBar1.Value & " / " & ProgressBar1.Max
                    'ProgressBar1.Value = ProgressBar1.Value
                    
                    'Set rsLite = conLite.OpenRecordset(strSQl, True)
                    cmd.SetText 1, rsAtt!employee_code
                    cmd.SetDate 2, rsAtt!att_date
                    cmd.SetText 3, IIf(IsNull(rsAtt!ip_address) = True, "", rsAtt!ip_address)
                    cmd.SetInt32 4, IIf(IsNull(rsAtt!enrollnumber) = True, 0, rsAtt!enrollnumber)
                    cmd.SetText 5, IIf(IsNull(rsAtt!shift_number) = True, 0, rsAtt!shift_number)
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
                    cmd.SetDate 29, IIf(IsNull(rsAtt!edit_date) = True, 0, rsAtt!edit_date)
                    cmd.SetText 30, IIf(IsNull(rsAtt!userinput) = True, 0, rsAtt!userinput)
                    cmd.SetText 31, IIf(IsNull(rsAtt!useredit) = True, 0, rsAtt!useredit)
                    cmd.Execute
                'End If
'            Next
                rsAtt.MoveNext
                Wend
            End If
            conLite.CommitTrans
        'End If
    
    MsgBox "Successfully Data Transferred. Total data transfered = " & ProgressBar1.Value & " !"
    ProgressBar1.Visible = False
    Label2.Visible = False

End Sub

Private Sub importDataEmployee()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset

    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("m_employee") Then
        MsgBox "Cannot find data on selected file.. Please check your file again !"
        Exit Sub
    Else
        strsql = "SELECT * FROM m_employee"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
    End If
        
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from m_employee", CnG, adOpenKeyset, adLockOptimistic
    
    'MsgBox rs

    For aa = 0 To rss.RecordCount - 1
        
        If Not check_exist_new_employee(rss!employee_code) Then
            i = i + 1
                
            With rs
                .AddNew
                                
                .Fields("employee_code").Value = rss!employee_code
                '-----------------------------------------------------------------------------
                .Fields("employee_name").Value = rss!employee_name
                .Fields("employee_nick_name").Value = rss!employee_nick_name
                
                .Fields("company_code").Value = rss!COMPANY_CODE
                .Fields("company_name").Value = rss!company_name
                .Fields("company_name").Value = rss!company_name
                .Fields("department_code").Value = rss!department_code
                .Fields("department_name").Value = rss!department_name
                .Fields("division_code").Value = rss!division_code
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
                .Fields("flag_shiftable").Value = rss!flag_shiftable
                .Fields("description").Value = rss!Description
                
                .Fields("flag_active").Value = rss!flag_active
            
                .Fields("end_working").Value = IIf(rss!end_working = "", "00:00:00", Format(rss!last_education_pass, "yyyy-mm-dd HH:mm:ss"))
                .Fields("reason").Value = rss!reason
                
                .Fields("fathers_name").Value = rss!fathers_name
                .Fields("mothers_name").Value = rss!mothers_name
                .Fields("child_number").Value = rss!child_number
                .Fields("child_number_from").Value = rss!child_number_from
                .Fields("managerial_level").Value = rss!managerial_level
                
                .Fields("start_mc").Value = IIf(rss!start_mc = "", "00:00:00", Format(rss!last_education_pass, "yyyy-mm-dd HH:mm:ss"))
                .Fields("end_mc").Value = IIf(rss!end_mc = "", "00:00:00", Format(rss!last_education_pass, "yyyy-mm-dd HH:mm:ss"))
                .Fields("branches").Value = rss!branches
                
                .Update
            End With
        End If
        rss.MoveNext
    Next
    
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
End Sub

Private Sub importDataAttendance()
Dim nmFile As String
Dim cmd As cCommand
Dim aa, bb, i As Integer
Dim rss As New cRecordset


    conLite.OpenDB txtFileName.Text
    
    If Not SelectTable("h_attendance") Then
        MsgBox "Cannot find data on selected file.. Please check your file again !"
        Exit Sub
    Else
        strsql = "SELECT * FROM h_attendance"
        rss.OpenRecordset strsql, conLite, True
        Set rsLite = conLite.OpenRecordset(strsql, True)
    End If
    
'    strsql = "SELECT * FROM h_attendance"
'
'    rss.OpenRecordset strsql, conLite, True
    
    'Set rsLite = conLite.OpenRecordset(strsql, True)

    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
    rs.Open "select * from h_attendance", CnG, adOpenKeyset, adLockOptimistic
                
    For aa = 0 To rss.RecordCount - 1
    
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
                .Fields("edit_date").Value = IIf(rss!edit_date = "", "00:00:00", Format(rss!edit_date, "yyyy-mm-dd HH:mm:ss"))
                .Fields("userinput").Value = rss!userinput
                .Fields("useredit").Value = rss!useredit
                '-----------------------
                .Update
            End With
        End If
        rss.MoveNext
    Next
    
    CnG.CommitTrans
    Else
        aa = 0
    End If
    
    MsgBox "Successfully Data Transferred. Total data transfered = " & i & " !"
    
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
'MsgBox "Cannot find data on selected file.. Please check your file again !"
SelectTable = False

End Function


