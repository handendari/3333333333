VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_import_salary 
   Caption         =   "Import Salary Data"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10785
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   5100
      TabIndex        =   4
      Top             =   9480
      Visible         =   0   'False
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   870
      Top             =   90
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
      MICON           =   "frm_import_salary.frx":0000
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
      Height          =   7980
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   20025
      Begin prj_fej_jkt.LynxGrid LynxGrid1 
         Height          =   7635
         Left            =   180
         TabIndex        =   3
         Top             =   210
         Width           =   15765
         _ExtentX        =   27808
         _ExtentY        =   13467
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
      MICON           =   "frm_import_salary.frx":001C
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
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   5130
      TabIndex        =   5
      Top             =   9780
      Visible         =   0   'False
      Width           =   7275
   End
End
Attribute VB_Name = "frm_import_salary"
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
        .AddColumn "month", 1200, lgAlignCenterCenter
        .AddColumn "employee code", 1500, lgAlignCenterCenter
        .AddColumn "date from", 1200, lgAlignCenterCenter, lgDate, "yyyy-MM-dd"
        .AddColumn "date to", 1200, lgAlignCenterCenter, lgDate, "yyyy-MM-dd"
        .AddColumn "salary basic", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "rate hour", 1000, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "allowance rate", 1000, lgAlignRightCenter, lgNumeric, "#,##" ' `allowance_rate` decimal(10,2) DEFAULT NULL,
        .AddColumn "meal allowance", 1000, lgAlignRightCenter, lgNumeric, "#,##" ' `uang_makan` decimal(10,2) DEFAULT NULL,
        .AddColumn "tot hour ot", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `jml_jam_lembur` decimal(10,2) DEFAULT NULL,
        .AddColumn "total OT", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `jml_lembur` decimal(15,2) DEFAULT NULL,
        .AddColumn "tot absent day", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' 'jml_hari_absen` int(11) DEFAULT NULL,
        .AddColumn "description", 1200 ' `description` varchar(50) DEFAULT NULL,
        .AddColumn "site allowance", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `tunj_lap` decimal(15,2) DEFAULT NULL,
        .AddColumn "insentive hadir", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `insentive_hadir` decimal(15,2) DEFAULT NULL,
        .AddColumn "other income", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `krg_bln_lalu` decimal(15,2) DEFAULT NULL,
        .AddColumn "thr", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `thr` decimal(15,2) DEFAULT NULL,
        .AddColumn "pot. absensi", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pot_absensi` decimal(15,2) DEFAULT NULL,
        .AddColumn "pot. others", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pot_others` decimal(15,2) DEFAULT NULL,
        .AddColumn "gaji kotor", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `gaji_kotor` decimal(15,2) DEFAULT NULL,
        .AddColumn "pot. pinjaman", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pot_pinjaman` decimal(15,2) DEFAULT NULL,
        .AddColumn "pot jamsostek", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pot_jamsostek` decimal(15,2) DEFAULT NULL,
        .AddColumn "pjk jkk jkm", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pjk_jkk_jkm_204` decimal(15,2) DEFAULT NULL,
        .AddColumn "tot. pendapatan kotor", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `tot_dpt_kotor` decimal(15,2) DEFAULT NULL,
        .AddColumn "pjk tunj_jabatan", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pjk_tunj_jabatan` decimal(15,2) DEFAULT NULL,
        .AddColumn "pjk ptkp", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pjk_ptkp` decimal(15,2) DEFAULT NULL,
        .AddColumn "pengurang pjk", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `tot_pengurang_pjk` decimal(15,2) DEFAULT NULL,
        .AddColumn "pendapatan kena pajak", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `tot_dpt_kena_pajak` decimal(15,2) DEFAULT NULL,
        .AddColumn "pendapatan kena pajak setahun", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' tot_dpt_kena_pajak_tahun` decimal(15,2) DEFAULT NULL,
        .AddColumn "pph 5 persen", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pph_5_persen` decimal(15,2) DEFAULT NULL,
        .AddColumn "pph 15 persen", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pph_15_persen` decimal(15,2) DEFAULT NULL,
        .AddColumn "pph 25 persen", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pph_25_persen` decimal(15,2) DEFAULT NULL,
        .AddColumn "pph 30 persen", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pph_30_persen` decimal(15,2) DEFAULT NULL,
        .AddColumn "pph21", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `pph21` decimal(15,2) DEFAULT NULL,
        .AddColumn "rounding", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `round` decimal(15,2) DEFAULT NULL,
        .AddColumn "gaji bersih", 1200, lgAlignRightCenter, lgNumeric, "#,##" ' `gaji_bersih` decimal(15,2) DEFAULT NULL,
        .BackColorBkg = &HFCE1CB
        .Redraw = True
   End With
    
End Sub

Private Sub Form_Load()
    createGrid
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Width = Me.Width - 500
    Frame1.Height = Me.Height - 1700
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
    Dim strWorksheet As String
    'Screen.MousePointer = vbHourglass
    'DoEvents
    strWorksheet = "data_salary"
    
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
                    LynxGrid1.AddItem .Fields("Month") & vbTab & .Fields("employee_code") & vbTab & .Fields("date_from") _
                            & vbTab & .Fields("date_to") & vbTab & .Fields("salary_basic") & vbTab & .Fields("rate_hour") _
                            & vbTab & .Fields("allowance_rate") & vbTab & .Fields("uang_makan") & vbTab & .Fields("jml_jam_lembur") _
                            & vbTab & .Fields("jml_lembur") & vbTab & .Fields("jml_hari_absen") & vbTab & .Fields("Description") _
                            & vbTab & .Fields("tunj_lap") & vbTab & .Fields("insentive_hadir") & vbTab & .Fields("krg_bln_lalu") _
                            & vbTab & .Fields("thr") & vbTab & .Fields("pot_absensi") & vbTab & .Fields("pot_others") & vbTab & .Fields("gaji_kotor") _
                            & vbTab & .Fields("pot_pinjaman") & vbTab & .Fields("pot_jamsostek") & vbTab & .Fields("pjk_jkk_jkm_204") _
                            & vbTab & .Fields("tot_dpt_kotor") & vbTab & .Fields("pjk_tunj_jabatan") & vbTab & .Fields("pjk_ptkp") & vbTab & .Fields("tot_pengurang_pjk") _
                            & vbTab & .Fields("tot_dpt_kena_pajak") & vbTab & .Fields("tot_dpt_kena_pajak_tahun") & vbTab & .Fields("pph_5_persen") & vbTab & .Fields("pph_15_persen") _
                            & vbTab & .Fields("pph_25_persen") & vbTab & .Fields("pph_30_persen") & vbTab & .Fields("pph21") & vbTab & .Fields("Round") & vbTab & .Fields("gaji_bersih")
                'End If
                .MoveNext
            Wend
            Me.MousePointer = vbNormal
        End If
    End With
    LynxGrid1.Redraw = True
End Sub

Private Sub vbButton2_Click()
Dim aa As Integer
    With LynxGrid1
        If .Rows > 0 Then
            ProgressBar1.Visible = True
            Label1.Visible = True
            ProgressBar1.Max = .Rows
            ProgressBar1.Value = 0
            For aa = 0 To .Rows - 1
                ProgressBar1.Value = aa
                Label1.Caption = .CellText(aa, 1) & " - " & .CellText(aa, 0)
                
                If .CellText(aa, 0) <> "" And .CellText(aa, 1) <> "" Then
                    strsql = "INSERT INTO h_salary_new VALUES " & _
                        "('" & .CellText(aa, 0) & "','" & .CellText(aa, 1) & "','" & Format(.CellText(aa, 2), "yyyy-MM-dd") & "','" & Format(.CellText(aa, 3), "yyyy-MM-dd") & "'," & _
                        "'" & IIf(IsNull(.CellValue(aa, 4)), 0, .CellValue(aa, 1)) & "','" & IIf(IsNull(.CellValue(aa, 5)), 0, .CellValue(aa, 5)) & "','" & IIf(IsNull(.CellValue(aa, 6)), 0, .CellValue(aa, 6)) & "','" & IIf(IsNull(.CellValue(aa, 7)), 0, .CellValue(aa, 7)) & "','" & IIf(IsNull(.CellValue(aa, 8)), 0, .CellValue(aa, 8)) & "'," & _
                        "'" & IIf(IsNull(.CellValue(aa, 9)), 0, .CellValue(aa, 9)) & "','" & IIf(IsNull(.CellValue(aa, 10)), 0, .CellValue(aa, 10)) & "','" & IIf(IsNull(.CellValue(aa, 11)), 0, .CellValue(aa, 11)) & "','" & IIf(IsNull(.CellValue(aa, 12)), 0, .CellValue(aa, 12)) & "','" & IIf(IsNull(.CellValue(aa, 13)), 0, .CellValue(aa, 13)) & "'," & _
                        "'" & IIf(IsNull(.CellValue(aa, 14)), 0, .CellValue(aa, 14)) & "','" & IIf(IsNull(.CellValue(aa, 15)), 0, .CellValue(aa, 15)) & "','" & IIf(IsNull(.CellValue(aa, 16)), 0, .CellValue(aa, 16)) & "','" & IIf(IsNull(.CellValue(aa, 17)), 0, .CellValue(aa, 17)) & "','" & IIf(IsNull(.CellValue(aa, 18)), 0, .CellValue(aa, 18)) & "'," & _
                        "'" & IIf(IsNull(.CellValue(aa, 19)), 0, .CellValue(aa, 19)) & "','" & IIf(IsNull(.CellValue(aa, 20)), 0, .CellValue(aa, 20)) & "','" & IIf(IsNull(.CellValue(aa, 21)), 0, .CellValue(aa, 21)) & "','" & IIf(IsNull(.CellValue(aa, 22)), 0, .CellValue(aa, 22)) & "','" & IIf(IsNull(.CellValue(aa, 23)), 0, .CellValue(aa, 23)) & "'," & _
                        "'" & IIf(IsNull(.CellValue(aa, 24)), 0, .CellValue(aa, 24)) & "','" & IIf(IsNull(.CellValue(aa, 25)), 0, .CellValue(aa, 25)) & "','" & IIf(IsNull(.CellValue(aa, 26)), 0, .CellValue(aa, 26)) & "','" & IIf(IsNull(.CellValue(aa, 27)), 0, .CellValue(aa, 27)) & "','" & IIf(IsNull(.CellValue(aa, 28)), 0, .CellValue(aa, 28)) & "'," & _
                        "'" & IIf(IsNull(.CellValue(aa, 29)), 0, .CellValue(aa, 29)) & "','" & IIf(IsNull(.CellValue(aa, 30)), 0, .CellValue(aa, 30)) & "','" & IIf(IsNull(.CellValue(aa, 31)), 0, .CellValue(aa, 31)) & "','" & IIf(IsNull(.CellValue(aa, 32)), 0, .CellValue(aa, 32)) & "','" & IIf(IsNull(.CellValue(aa, 33)), 0, .CellValue(aa, 33)) & "'," & _
                        "'" & IIf(IsNull(.CellValue(aa, 34)), 0, .CellValue(aa, 34)) & "')"
                    CnG.Execute strsql
                End If
                DoEvents
            Next
            MsgBox "Import Data Success...!!!"
            
            ProgressBar1.Visible = False
            Label1.Visible = False
        End If
    End With
End Sub
