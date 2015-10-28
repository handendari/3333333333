VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_import_salary_pph 
   Caption         =   "Import Salary Data"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10785
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1230
      TabIndex        =   6
      Top             =   90
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   4260
      TabIndex        =   4
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
   Begin prj_fej_jkt.vbButton vbButton1 
      Height          =   555
      Left            =   450
      TabIndex        =   1
      Top             =   9510
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      BTYPE           =   14
      TX              =   "Browse File"
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
      MICON           =   "frm_import_salary_pph.frx":0000
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
      Height          =   8400
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   20025
      Begin prj_fej_jkt.LynxGrid LynxGrid1 
         Height          =   8025
         Left            =   180
         TabIndex        =   3
         Top             =   210
         Width           =   15765
         _ExtentX        =   27808
         _ExtentY        =   14155
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
   Begin prj_fej_jkt.vbButton vbButton2 
      Height          =   555
      Left            =   2070
      TabIndex        =   2
      Top             =   9510
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      BTYPE           =   14
      TX              =   "Import Data"
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
      MICON           =   "frm_import_salary_pph.frx":001C
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
   Begin prj_fej_jkt.vbButton vbButton3 
      Height          =   555
      Left            =   14280
      TabIndex        =   8
      Top             =   9540
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      BTYPE           =   14
      TX              =   "Exit"
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
      MICON           =   "frm_import_salary_pph.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   7
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4290
      TabIndex        =   5
      Top             =   9780
      Visible         =   0   'False
      Width           =   7275
   End
End
Attribute VB_Name = "frm_import_salary_pph"
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
        .AddColumn "employee code", 1500, lgAlignCenterCenter, lgString
        .AddColumn "employee name", 2000, lgAlignLeftCenter
        .AddColumn "npwp", 1200, lgAlignCenterCenter
        .AddColumn "birth place", 1500, lgAlignLeftCenter
        .AddColumn "birth date", 1200, lgAlignCenterCenter, lgDate, "yyyy-MM-dd"
        .AddColumn "address", 3000, lgAlignLeftCenter
        .AddColumn "department", 1500, lgAlignCenterCenter
        .AddColumn "title", 1500, lgAlignCenterCenter
        .AddColumn "location", 1500, lgAlignCenterCenter
        .AddColumn "marital status", 1000, lgAlignCenterCenter
        .AddColumn "sex", 500, lgAlignCenterCenter
        .AddColumn "employement date", 1200, lgAlignCenterCenter, lgDate, "yyyy-MM-dd"
        .AddColumn "basic salary", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "overtime", 1000, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "site allowance", 1000, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "insentive", 1000, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "income others", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "bonus / thr", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "total gross", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "jht", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "pph21", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "absensi", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "deduction other", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "total deduction", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "net salary", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "loan", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "round", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "tot. payment", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .BackColorBkg = &HFCE1CB
        .Redraw = True
   End With
    
End Sub

Private Sub Form_Load()
    createGrid
    
    txtYear.Text = Year(Now())
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Width = Me.Width - 500
    Frame1.Height = Me.Height - 2200
    LynxGrid1.Height = Frame1.Height - 400
    LynxGrid1.Width = Frame1.Width - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_import_salary = Nothing
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
On Error GoTo err
    Dim strWorksheet As String
    'Screen.MousePointer = vbHourglass
    'DoEvents
    strWorksheet = "data_salary_eks"
    
    Adodc1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & str_file_name & ";Extended Properties=Excel 8.0"
    
    Adodc1.RecordSource = "select * from [" & strWorksheet & "$] order by employee_code asc"
    Adodc1.Refresh
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    With Adodc1.Recordset
        If .RecordCount > 0 Then
            Me.MousePointer = vbHourglass
            .MoveFirst
            While Not .EOF
                'If Adodc1.Recordset!employee_code <> "" Or Adodc1.Recordset!employee_code Is Null Then
                    LynxGrid1.AddItem .Fields("employee_code") & vbTab & .Fields("employee_name") & vbTab & .Fields("npwp") _
                            & vbTab & .Fields("birth_place") & vbTab & .Fields("birth_date") & vbTab & .Fields("address") _
                            & vbTab & .Fields("department") & vbTab & .Fields("title") & vbTab & .Fields("location") _
                            & vbTab & .Fields("marital_status") & vbTab & .Fields("sex") & vbTab & .Fields("employment_date") _
                            & vbTab & .Fields("basic_salary") & vbTab & .Fields("overtime") & vbTab & .Fields("site_allowance") _
                            & vbTab & .Fields("incentive") & vbTab & .Fields("others") & vbTab & .Fields("bonus") & vbTab & .Fields("tot_gross") _
                            & vbTab & .Fields("ded_jht") & vbTab & .Fields("ded_pph21") & vbTab & .Fields("ded_absen") _
                            & vbTab & .Fields("ded_other") & vbTab & .Fields("tot_ded") & vbTab & .Fields("net_salary") & vbTab & .Fields("ded_loan") _
                            & vbTab & .Fields("round") & vbTab & .Fields("tot_payment")
                'End If
                .MoveNext
            Wend
            Me.MousePointer = vbNormal
        End If
    End With
    LynxGrid1.Redraw = True

err:
If err.Number = "-2147467259" Then
    MsgBox "Your File Already Open or Invalid Sheet Name!", vbExclamation, headerMSG
    Exit Sub
End If
End Sub

Private Sub vbButton2_Click()
Dim aa As Integer

On Error GoTo err

If txtYear.Text = "" Then
    MsgBox "Please Fill Year First!", vbExclamation, headerMSG
    Exit Sub
End If

strsql = "DELETE from h_salary_eks WHERE year = '" & txtYear.Text & "'"
CnG.Execute strsql

With LynxGrid1
    If .Rows > 0 Then
        ProgressBar1.Visible = True
        Label1.Visible = True
        ProgressBar1.Max = .Rows
        ProgressBar1.Value = 0
        i = 0
        For aa = 0 To .Rows - 1
            ProgressBar1.Value = aa
            Label1.Caption = .CellText(aa, 1) & " - " & .CellText(aa, 0)
            
            If .CellText(aa, 0) <> "" And .CellText(aa, 1) <> "" Then
                strsql = "INSERT INTO h_salary_eks VALUES " & _
                    "('" & txtYear.Text & "','" & .CellText(aa, 0) & "','" & Replace(.CellText(aa, 1), "'", "") & "','" & IIf(IsNull(.CellText(aa, 2)), 0, .CellText(aa, 2)) & "','" & IIf(IsNull(.CellText(aa, 3)), 0, .CellText(aa, 3)) & "'," & _
                    "'" & Format(.CellText(aa, 4), "yyyy-MM-dd") & "','" & Replace(IIf(IsNull(.CellText(aa, 5)), 0, .CellText(aa, 5)), "'", "") & "','" & IIf(IsNull(.CellText(aa, 6)), 0, .CellText(aa, 6)) & "','" & IIf(IsNull(.CellText(aa, 7)), 0, .CellText(aa, 7)) & "','" & IIf(IsNull(.CellText(aa, 8)), 0, .CellText(aa, 8)) & "'," & _
                    "'" & IIf(IsNull(.CellText(aa, 9)), 0, .CellText(aa, 9)) & "','" & IIf(IsNull(.CellText(aa, 10)), 0, .CellText(aa, 10)) & "','" & Format(.CellText(aa, 11), "yyyy-MM-dd") & "','" & IIf(IsNull(.CellValue(aa, 12)), 0, .CellValue(aa, 12)) & "','" & IIf(IsNull(.CellValue(aa, 13)), 0, .CellValue(aa, 13)) & "'," & _
                    "'" & IIf(IsNull(.CellValue(aa, 14)), 0, .CellValue(aa, 14)) & "','" & IIf(IsNull(.CellValue(aa, 15)), 0, .CellValue(aa, 15)) & "','" & IIf(IsNull(.CellValue(aa, 16)), 0, .CellValue(aa, 16)) & "','" & IIf(IsNull(.CellValue(aa, 17)), 0, .CellValue(aa, 17)) & "','" & IIf(IsNull(.CellValue(aa, 18)), 0, .CellValue(aa, 18)) & "'," & _
                    "'" & IIf(IsNull(.CellValue(aa, 19)), 0, .CellValue(aa, 19)) & "','" & IIf(IsNull(.CellValue(aa, 20)), 0, .CellValue(aa, 20)) & "','" & IIf(IsNull(.CellValue(aa, 21)), 0, .CellValue(aa, 21)) & "','" & IIf(IsNull(.CellValue(aa, 22)), 0, .CellValue(aa, 22)) & "','" & IIf(IsNull(.CellValue(aa, 23)), 0, .CellValue(aa, 23)) & "'," & _
                    "'" & IIf(IsNull(.CellValue(aa, 24)), 0, .CellValue(aa, 24)) & "','" & IIf(IsNull(.CellValue(aa, 25)), 0, .CellValue(aa, 25)) & "','" & IIf(IsNull(.CellValue(aa, 26)), 0, .CellValue(aa, 26)) & "','" & IIf(IsNull(.CellValue(aa, 27)), 0, .CellValue(aa, 27)) & "')"
                CnG.Execute strsql
            End If
            DoEvents
        Next
        MsgBox "Import Data Success with " & i & " data Skipped"
        
        ProgressBar1.Visible = False
        Label1.Visible = False
    End If
End With

err:
If err.Number = "-2147217900" Then
    MsgBox "Employee Code " & "'" & LynxGrid1.CellText(aa, 0) & "'" & " Skipped Because There is Same!" & Chr(13) & _
        "Please Check Your Data.", vbExclamation, headerMSG
        i = i + 1
    Resume Next
End If

End Sub

Private Sub vbButton3_Click()
Unload Me
End Sub
