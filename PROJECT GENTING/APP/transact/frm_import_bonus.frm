VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_trans_import_bonus 
   Caption         =   "Import Bonus"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10785
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin prj_genting.vbButton vbButton2 
      Height          =   555
      Left            =   2070
      TabIndex        =   5
      Top             =   9510
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   979
      BTYPE           =   14
      TX              =   "&Import"
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
      MICON           =   "frm_import_bonus.frx":0000
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
      Height          =   555
      Left            =   390
      TabIndex        =   4
      Top             =   9510
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   979
      BTYPE           =   14
      TX              =   "&Browse"
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
      MICON           =   "frm_import_bonus.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6750
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   4260
      TabIndex        =   1
      Top             =   9480
      Visible         =   0   'False
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7260
      Top             =   150
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.Frame Frame1 
      Height          =   8190
      Left            =   60
      TabIndex        =   0
      Top             =   750
      Width           =   20025
      Begin prj_genting.LynxGrid LynxGrid1 
         Height          =   8175
         Left            =   210
         TabIndex        =   3
         Top             =   300
         Width           =   18375
         _ExtentX        =   32411
         _ExtentY        =   14420
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
   Begin prj_genting.vbButton vbButton3 
      Height          =   555
      Left            =   14370
      TabIndex        =   6
      Top             =   9540
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   979
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
      MICON           =   "frm_import_bonus.frx":0038
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
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "IMPORT BONUS"
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
      Left            =   6705
      TabIndex        =   7
      Top             =   30
      Width           =   2445
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4290
      TabIndex        =   2
      Top             =   9780
      Visible         =   0   'False
      Width           =   7275
   End
End
Attribute VB_Name = "frm_trans_import_bonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset
Dim strsql As String

Private Sub DTPicker1_Validate(Cancel As Boolean)
    LynxGrid1.Clear
End Sub

Private Sub createGrid()
   With LynxGrid1
        .AddColumn "Tanggal", 1700, lgAlignCenterCenter, lgDate, "yyyy-MM-dd"
        .AddColumn "Kode Karyawan", 1500, lgAlignCenterCenter, lgString
        .AddColumn "Kode Prestasi", 1700, lgAlignLeftCenter, lgString
        .AddColumn "Kali Gaji", 1000, lgAlignCenterCenter, lgNumeric, "#,##"
        .AddColumn "Keterangan", 1700, lgAlignLeftCenter, lgString
        .BackColorBkg = &HFCE1CB
        .Redraw = True
   End With
    
End Sub

Private Sub Form_Load()
    createGrid
End Sub

Private Sub Form_Resize()
    Frame1.Width = Me.Width - 500
    Frame1.Height = Me.Height - 2200
    LynxGrid1.Height = Frame1.Height - 400
    LynxGrid1.Width = Frame1.Width - 400
'    vbButton1.Top = LynxGrid1.Top + LynxGrid1.Height + 700
'    vbButton2.Top = LynxGrid1.Top + LynxGrid1.Height + 700
'    vbButton3.Top = LynxGrid1.Top + LynxGrid1.Height + 700
'    ProgressBar1.Top = LynxGrid1.Top + LynxGrid1.Height + 700
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_import_bonus = Nothing
End Sub

Private Sub vbButton1_Click()
    CommonDialog1.Filter = "XLS|*.xls"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
        Call fill_grid_excel_m(CommonDialog1.FileName)
    End If
End Sub

Private Sub fill_grid_excel_m(str_file_name As String)
'Dim rsemp As New ADODB.Recordset
Dim strsql As String

On Error GoTo Err
    Dim strWorksheet As String
    'Screen.MousePointer = vbHourglass
    'DoEvents
    strWorksheet = "data_bonus"
    
    Adodc1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & str_file_name & ";Extended Properties=Excel 8.0"
    
    Adodc1.RecordSource = "select * from [" & strWorksheet & "$]"
    Adodc1.Refresh
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    With Adodc1.Recordset
        If .RecordCount > 0 Then
            Me.MousePointer = vbHourglass
            .MoveFirst
            While Not .EOF
                'If Adodc1.Recordset!employee_code <> "" Or Adodc1.Recordset!employee_code Is Null Then
                LynxGrid1.AddItem .Fields(0) & vbTab & Trim(.Fields(1)) _
                    & vbTab & .Fields(2) & vbTab & .Fields(3) & vbTab & .Fields(4)
                'End If
                .MoveNext
            Wend
            Me.MousePointer = vbNormal
        End If
    End With
    LynxGrid1.Redraw = True

Err:
If Err.Number = "-2147467259" Then
    MsgBox "Your File Already Open or Invalid Sheet Name!", vbExclamation, headerMSG
    Exit Sub
End If
End Sub

Private Sub vbbutton2_Click()
Dim aa As Integer
Dim rs2 As New ADODB.Recordset
Dim start_time As String, end_time As String, max_break_out As String, min_break_in As String
Dim time_in As String, time_out_break As String, time_in_break As String, time_out As String

'On Error GoTo err
            
With LynxGrid1
        If .Rows > 0 Then
            ProgressBar1.Visible = True
            Label1.Visible = True
            ProgressBar1.Max = .Rows
            ProgressBar1.Value = 0
            For aa = 1 To .Rows
                ProgressBar1.Value = aa
                Label1.Caption = .CellText(aa, 0) & " - " & .CellText(aa, 1)
                
                Dim v_employee_code As Integer
                strsql = "SELECT employee_code FROM m_employee " _
                        & "WHERE nik = '" & .CellText(i, 1) & "'"
                rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    v_employee_code = rs!employee_code
                End If
                rs.Close
                
                strsql = "DELETE FROM t_employee_performance " _
                        & "WHERE performance_date = '" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "' AND " _
                        & "employee_code = '" & v_employee_code & "'"
                CnG.Execute strsql
                        
                If .CellText(aa, 0) <> "" And .CellText(aa, 1) <> "" Then
                    strsql = "INSERT INTO t_employee_performance(performance_date,employee_code," _
                        & "prestasi_code,kali_gaji,entry_date,entry_user,description) VALUES " _
                        & "('" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "'," _
                        & "'" & v_employee_code & "','" & .CellText(aa, 2) & "'," _
                        & "'" & .CellValue(aa, 3) & "',now(),'" & LOGIN_CODE & "'," _
                        & "'" & .CellText(aa, 4) & "')"
                    CnG.Execute strsql
                End If
                DoEvents
            Next
            MsgBox "Import Data Success...!!!"
            
            ProgressBar1.Visible = False
            Label1.Visible = False
        End If
    End With

'err:
'If err.Number = "-2147217900" Then
'    MsgBox "Employee Code " & "'" & LynxGrid1.CellText(aa, 0) & "'" & " Skipped Because There is Same!" & Chr(13) & _
'        "Please Check Your Data.", vbExclamation, headerMSG
'        i = i + 1
'    Resume Next
'End If

End Sub

Private Sub vbButton3_Click()
Unload Me
End Sub
