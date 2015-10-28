VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_export_pph 
   Caption         =   "EXPORT PPH21"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12270
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10785
   ScaleWidth      =   12270
   WindowState     =   2  'Maximized
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   0
      Top             =   0
   End
   Begin prj_tpc.LynxGrid LynxGrid1 
      Height          =   7065
      Left            =   210
      TabIndex        =   11
      Top             =   1740
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   12462
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
   Begin prj_tpc.vbButton vbButton1 
      Height          =   585
      Left            =   2160
      TabIndex        =   9
      Top             =   9540
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
      BTYPE           =   14
      TX              =   "Export ke CSV"
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
      MICON           =   "frm_export_pph.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton vbButton5 
      Height          =   585
      Left            =   540
      TabIndex        =   8
      Top             =   9540
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1032
      BTYPE           =   14
      TX              =   "Export ke Excel"
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
      MICON           =   "frm_export_pph.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton vbButton4 
      Height          =   315
      Left            =   10110
      TabIndex        =   7
      Top             =   1050
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Load"
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
      MICON           =   "frm_export_pph.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3210
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1050
      Width           =   3855
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8640
      TabIndex        =   2
      Top             =   1050
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   4260
      TabIndex        =   1
      Top             =   9660
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
      Left            =   120
      Top             =   1410
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
      Height          =   7440
      Left            =   60
      TabIndex        =   0
      Top             =   1500
      Width           =   20025
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   1290
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1470
      OleObjectBlob   =   "frm_export_pph.frx":0054
      TabIndex        =   5
      Top             =   1050
      Width           =   1695
   End
   Begin prj_tpc.vbButton vbButton2 
      Height          =   585
      Left            =   4260
      TabIndex        =   10
      Top             =   9930
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
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
      MICON           =   "frm_export_pph.frx":1FBA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton vbButton3 
      Height          =   585
      Left            =   14160
      TabIndex        =   12
      Top             =   9510
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
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
      MICON           =   "frm_export_pph.frx":1FD6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "EXPORT PPH21"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Top             =   210
      Width           =   4365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   510
      TabIndex        =   6
      Top             =   1110
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7680
      TabIndex        =   3
      Top             =   1110
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_export_pph.frx":1FF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21450
   End
End
Attribute VB_Name = "frm_export_pph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset

Private Sub DTPicker1_Validate(Cancel As Boolean)
    LynxGrid1.Clear
End Sub

Private Sub createGrid()
   With LynxGrid1
        .AddColumn "Kode Form", 1500, lgAlignCenterCenter, lgString
        .AddColumn "Tahun Pajak", 700, lgAlignCenterCenter
        .AddColumn "Pembetulan", 700, lgAlignCenterCenter
        .AddColumn "Nomor Urut", 700, lgAlignCenterCenter
        .AddColumn "NPWP Pegawai", 1500, lgAlignLeftCenter
        .AddColumn "Nama Pegawai", 1500, lgAlignLeftCenter
        .AddColumn "Alamat Pegawai", 3000, lgAlignLeftCenter
        .AddColumn "Jabatan Pegawai", 1500, lgAlignLeftCenter
        .AddColumn "Jenis Kelamin", 700, lgAlignCenterCenter
        .AddColumn "Status Pegawai", 700, lgAlignCenterCenter
        .AddColumn "Status Kawin", 700, lgAlignCenterCenter
        .AddColumn "FLG_ASING", 700, lgAlignCenterCenter
        .AddColumn "Status PTKP", 700, lgAlignCenterCenter
        .AddColumn "Jumlah Tanggungan", 700, lgAlignCenterCenter
        .AddColumn "Masa Perolehan 1", 700, lgAlignCenterCenter
        .AddColumn "Masa Perolehan 2", 700, lgAlignCenterCenter
        .AddColumn "A1", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "FLG_A2", 700, lgAlignCenterCenter
        .AddColumn "A2", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A3", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A4", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A5", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A6", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A7", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A8", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A9", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A10", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A11", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A12", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A13", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A14", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A15", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A16", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A17", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A18", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A19", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A20", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A21", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A22", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A22a", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A22b", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A23", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "A24", 1200, lgAlignRightCenter, lgNumeric, "#,##"
        .AddColumn "FLG_A24", 1200, lgAlignCenterCenter
        .AddColumn "BLN_A24", 1200, lgAlignCenterCenter
        .BackColorBkg = &HFCE1CB
        .Redraw = True
   End With
    
End Sub

Private Sub Form_Load()
    Call load_data_company
    createGrid
    
    txtYear.Text = year(Now())
    
    timer1.Enabled = True
    
    vbButton1.Enabled = False
    vbButton5.Enabled = False
End Sub

Private Sub Form_Resize()
    Frame1.Width = Me.Width - 500
    'Frame1.Height = Me.Height - 2200
    LynxGrid1.Height = Frame1.Height - 400
    LynxGrid1.Width = Frame1.Width - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_export_pph = Nothing
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If
    
    vbButton1.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)
    vbButton5.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)
End Sub

Private Sub timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub vbButton5_Click()
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

On Error GoTo Err

    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
   
Dim DataArray(1 To 200, 1 To 45) As Variant
Dim r As Integer
Dim NumberOfRows As Integer
Dim v_loop As Integer
Dim nmFile As String
        
        CommonDialog1.Filter = "Excel 97-2003 Workbook (*.xls)"
        CommonDialog1.ShowSave
        
        If Len(CommonDialog1.FileName) <> 0 Then
            If Right(CommonDialog1.FileName, 4) = ".xls" Then
                nmFile = CommonDialog1.FileName
            Else
                nmFile = CommonDialog1.FileName & ".xls"
            End If
        Else
            Exit Sub
        End If
        
'        If checkFile(nmFile) Then
'        a = MsgBox("File has been existed ! " & _
'            "Overwrite File ?", vbOKCancel, headerMSG)
'            If a <> vbOK Then
'                rs.Close
'                Exit Sub
'            End If
'        End If
        
    SQL = "SELECT  kode_form,tahun_pajak,pembetulan,nomor_urut,npwp_pegawai, " _
            & "nama_pegawai,alamat_pegawai,jabatan_pegawai,jenis_kelamin,status_pegawai, " _
            & "status_kawin,flg_asing,status_ptkp,jumlah_tanggungan,masa_perolehan_1, " _
            & "masa_perolehan_2,a1,flg_a2,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13, " _
            & "a14,a15,a16,a17,a18,a19,a20,a21,a22,a22a,a22b,a23,a24,flg_a24,bln_a24 " _
        & "from t_pph_export " _
        & "where tahun_pajak = '" & txtYear.Text & "' AND company_code = '" & TDBCombo_company.Text & "' " _
        & "ORDER by kode_form"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    ProgressBar1.Visible = True
    ProgressBar1.Max = rs.RecordCount
    ProgressBar1.Value = 0
    
    NumberOfRows = rs.RecordCount
    rs.MoveFirst
    
    For r = 1 To NumberOfRows
    ProgressBar1.Value = r
    
    Dim vRecord As String
        For v_loop = 1 To 45
        v_record = IIf(v_loop = 1, rs!kode_form, IIf(v_loop = 2, rs!tahun_pajak, IIf(v_loop = 3, rs!pembetulan, _
                IIf(v_loop = 4, rs!nomor_urut, IIf(v_loop = 5, rs!npwp_pegawai, IIf(v_loop = 6, rs!nama_pegawai, _
                IIf(v_loop = 7, rs!alamat_pegawai, IIf(v_loop = 8, rs!jabatan_pegawai, IIf(v_loop = 9, rs!jenis_kelamin, _
                IIf(v_loop = 10, rs!status_pegawai, IIf(v_loop = 11, rs!status_kawin, IIf(v_loop = 12, rs!flg_asing, _
                IIf(v_loop = 13, rs!status_ptkp, IIf(v_loop = 14, rs!jumlah_tanggungan, IIf(v_loop = 15, rs!masa_perolehan_1, _
                IIf(v_loop = 16, rs!masa_perolehan_2, IIf(v_loop = 17, rs!A1, IIf(v_loop = 18, rs!flg_a2, IIf(v_loop = 19, rs!A2, _
                IIf(v_loop = 20, rs!A3, IIf(v_loop = 21, rs!a4, IIf(v_loop = 22, rs!a5, IIf(v_loop = 23, rs!a6, IIf(v_loop = 24, rs!a7, IIf(v_loop = 25, rs!a8, _
                IIf(v_loop = 26, rs!a9, IIf(v_loop = 27, rs!a10, IIf(v_loop = 28, rs!a11, IIf(v_loop = 29, rs!a12, IIf(v_loop = 30, rs!a13, _
                IIf(v_loop = 31, rs!a14, IIf(v_loop = 32, rs!a15, IIf(v_loop = 33, rs!a16, IIf(v_loop = 34, rs!a17, IIf(v_loop = 35, rs!a18, IIf(v_loop = 36, rs!a19, _
                IIf(v_loop = 37, rs!a20, IIf(v_loop = 38, rs!a21, IIf(v_loop = 39, rs!a22, IIf(v_loop = 40, rs!a22a, IIf(v_loop = 41, rs!a22b, _
                IIf(v_loop = 42, rs!a23, IIf(v_loop = 43, rs!a24, IIf(v_loop = 44, rs!flg_a24, rs!bln_a24))))))))))))))))))))))))))))))))))))))))))))
        
        DataArray(r, v_loop) = v_record
        Next v_loop

    rs.MoveNext
    Next
    Set oSheet = oBook.Worksheets(1)
   
    oSheet.Range("A1:AS1").Value = Array("Kode Form", "Tahun Pajak", "Pembetulan", "Nomor Urut", _
                                "NPWP Pegawai", "Nama Pegawai", "Alamat Pegawai", "Jabatan Pegawai", _
                                "Jenis Kelamin", "Status Pegawai", "Status Kawin", "FLG_ASING", _
                                "Status PTKP", "Jumlah Tanggungan", "Masa Perolehan 1", "Masa Perolehan 2", _
                                "A1", "FLG_A2", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", _
                                "A13", "A14", "A15", "A16", "A17", "A18", "A19", "A20", "A21", "A22", "A22a", "A22b", _
                                "A23", "A24", "FLG_A24", "BLN_A24")
            ' Put headers of fields to excel file
   
    oSheet.Range("A2").Resize(NumberOfRows, 45).Value = DataArray
   
    oBook.SaveAs nmFile
    oExcel.Quit
   
    rs.MoveFirst
    rs.Close
    
    ProgressBar1.Visible = False
    MsgBox "Export File to Excel Success..", vbInformation, headerMSG

Err:
If Err.Number = 1004 Then
    ProgressBar1.Visible = False
    MsgBox "Data Not Be Imported!", vbInformation, headerMSG
    rs.Close
    Exit Sub
End If
End Sub

Private Sub vbButton1_Click()

    SQL = "SELECT  kode_form,tahun_pajak,pembetulan,nomor_urut,npwp_pegawai, " _
                & "nama_pegawai,alamat_pegawai,jabatan_pegawai,jenis_kelamin,status_pegawai, " _
                & "status_kawin,flg_asing,status_ptkp,jumlah_tanggungan,masa_perolehan_1, " _
                & "masa_perolehan_2,a1,flg_a2,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13, " _
                & "a14,a15,a16,a17,a18,a19,a20,a21,a22,a22a,a22b,a23,a24,flg_a24,bln_a24 " _
            & "from t_pph_export " _
            & "where tahun_pajak = '" & txtYear.Text & "' AND company_code = '" & TDBCombo_company.Text & "' " _
            & "ORDER by kode_form"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    If rs.RecordCount > 0 Then
            Dim jmlrecord As Long
            Dim nmFile As String
            
            CommonDialog1.Filter = "CSV (Comma delimited) (*.csv)"
    '        CommonDialog1.InitDir = App.Path
            CommonDialog1.ShowSave
            
            If Len(CommonDialog1.FileName) <> 0 Then
                If Right(CommonDialog1.FileName, 4) = ".csv" Then
                    nmFile = CommonDialog1.FileName
                Else
                    nmFile = CommonDialog1.FileName & ".csv"
                End If
            Else
                rs.Close
                Exit Sub
            End If
            
            If checkFile(nmFile) Then
            a = MsgBox("File has been existed ! " & _
                "Overwrite File ?", vbOKCancel, headerMSG)
                If a <> vbOK Then
                    rs.Close
                    Exit Sub
                End If
    '        Else
    '            Open nmFile For Output Access Write As #1
            End If
            
            Open nmFile For Output Access Write As #1
            
            jmlrecord = rs.RecordCount
            'Label3.Caption = "0 of " & jmlrecord & " Records"
            ProgressBar1.Value = 0
            ProgressBar1.Max = jmlrecord
            ProgressBar1.Visible = True
        
            Dim i As Integer
            Dim fieldNumber As Integer
            Dim HeaderName As String
            With rs
                fieldNumber = .Fields.Count - 1
                For i = 0 To fieldNumber      'Now add the field names
                    Select Case i
                        Case 0
                            HeaderName = "Kode Form"
                        Case 1
                            HeaderName = "Tahun Pajak"
                        Case 2
                            HeaderName = "Pembetulan"
                        Case 3
                            HeaderName = "Nomor Urut"
                        Case 4
                            HeaderName = "NPWP Pegawai"
                        Case 5
                            HeaderName = "Nama Pegawai"
                        Case 6
                            HeaderName = "Alamat Pegawai"
                        Case 7
                            HeaderName = "Jabatan Pegawai"
                        Case 8
                            HeaderName = "Jenis Kelamin"
                        Case 9
                            HeaderName = "Status Pegawai"
                        Case 10
                            HeaderName = "Status Kawin"
                        Case 11
                            HeaderName = "FLG_ASING"
                        Case 12
                            HeaderName = "Status PTKP"
                        Case 13
                            HeaderName = "Jumlah Tanggungan"
                        Case 14
                            HeaderName = "Masa Perolehan 1"
                        Case 15
                            HeaderName = "Masa Perolehan 2"
                        Case 16
                            HeaderName = "A1"
                        Case 17
                            HeaderName = "FLG_A2"
                        Case 18
                            HeaderName = "A2"
                        Case 19
                            HeaderName = "A3"
                        Case 20
                            HeaderName = "A4"
                        Case 21
                            HeaderName = "A5"
                        Case 22
                            HeaderName = "A6"
                        Case 23
                            HeaderName = "A7"
                        Case 24
                            HeaderName = "A8"
                        Case 25
                            HeaderName = "A9"
                        Case 26
                            HeaderName = "A10"
                        Case 27
                            HeaderName = "A11"
                        Case 28
                            HeaderName = "A12"
                        Case 29
                            HeaderName = "A13"
                        Case 30
                            HeaderName = "A14"
                        Case 31
                            HeaderName = "A15"
                        Case 32
                            HeaderName = "A16"
                        Case 33
                            HeaderName = "A17"
                        Case 34
                            HeaderName = "A18"
                        Case 35
                            HeaderName = "A19"
                        Case 36
                            HeaderName = "A20"
                        Case 37
                            HeaderName = "A21"
                        Case 38
                            HeaderName = "A22"
                        Case 39
                            HeaderName = "A22a"
                        Case 40
                            HeaderName = "A22b"
                        Case 41
                            HeaderName = "A23"
                        Case 42
                            HeaderName = "A24"
                        Case 43
                            HeaderName = "FLG_A24"
                        Case 44
                            HeaderName = "BLN_A24"
                        End Select
                    Print #1, HeaderName & ";"; 'similar to the ones below
                Next i
                Print #1, ""
        
                Dim recKe As Long
                recKe = 0
                .MoveFirst
                While Not .EOF
                    recKe = recKe + 1
                    For i = 0 To fieldNumber      'If there is an emty field,
                        If (IsNull(.Fields(i))) Then    'add a , to indicate it is
                            Print #1, "|";              'empty
                        Else
                            If i = fieldNumber Then
                                Print #1, Trim$(CStr(.Fields(i)));
                            Else
                                Print #1, Trim$(CStr(.Fields(i))) & ";";
                            End If
                        End If                  'Putting data under "" will not
                    Next i                      'confuse the reader of the file
                    'DoEventsEx                  'between Dhaka, Bangladesh as two
                    Print #1,                   'fields or as one field.
                    'Label3.Caption = "Process " & recKe & " of " & jmlrecord & " Records"
                    ProgressBar1.Value = recKe
                    DoEvents
                    .MoveNext
                Wend
            End With
            Close #1
            
            ProgressBar1.Visible = False
            MsgBox "Export File to CSV Success..", vbInformation, headerMSG
    End If
    rs.Close
End Sub

Private Sub vbButton3_Click()
    Unload Me
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub vbButton4_Click()
Dim cls_insert_export_pph As New cls_insert_export_pph
Dim SQL As String
Dim rs As New Recordset

    SQL = "Select a.employee_code " _
            & "from h_salary a join m_employee b on a.employee_code = b.employee_code " _
            & "where left(a.month,4) = '" & txtYear.Text & "' AND b.company_code = '" & TDBCombo_company.Text & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        MsgBox "No Data Found With Branch Office " & TDBCombo_company.Text & " and Year " & txtYear & "!", vbExclamation, headerMSG
        rs.Close
        Exit Sub
    End If
    rs.Close
    
    ProgressBar1.Visible = True
            
    SQL = "Select Distinct(a.employee_code) " _
            & "from h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
            & "JOIN h_attendance c ON a.employee_code = c.employee_code " _
            & "Where left(month,4) = '" & txtYear.Text & "' AND b.company_code = '" & TDBCombo_company.Text & "' " _
            & "GROUP BY a.employee_code"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    CnG.BeginTrans
        
    If rs.RecordCount > 0 Then
        Dim rsnourut As New Recordset
        Dim nourut As Double
        Dim i As Double
        
        ProgressBar1.Max = rs.RecordCount
        ProgressBar1.Value = 0
        
        SQL = "DELETE FROM t_pph_export WHERE tahun_pajak = '" & txtYear.Text & "' " _
                & "AND company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute SQL
        
        SQL = "Select MAX(nourut) nourut from t_pph_export"
        rsnourut.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If Not rs.EOF Then
            nourut = IIf(IsNull(rsnourut!nourut), 0, rsnourut!nourut)
        End If
        
        i = nourut
        rs.MoveFirst
            While Not rs.EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            i = i + 1
        
                Call cls_insert_export_pph.insert_export_pph(txtYear, TDBCombo_company, rs!employee_code, i)
            
            rs.MoveNext
            Wend
    
        End If
        rs.Close
            
    CnG.CommitTrans
        
ProgressBar1.Visible = False
Call isiGridAbsen
End Sub

Public Sub isiGridAbsen()
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    'If rs.State = 1 Then rs.Close
    SQL = "select * from t_pph_export Where tahun_pajak = '" & txtYear & "' " _
            & "AND company_code = '" & TDBCombo_company & "' ORDER BY kode_form, tahun_pajak"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            LynxGrid1.AddItem rs!kode_form & vbTab & rs!tahun_pajak & vbTab & rs!pembetulan _
                & vbTab & rs!nomor_urut & vbTab & rs!npwp_pegawai & vbTab & rs!nama_pegawai _
                & vbTab & rs!alamat_pegawai & vbTab & rs!jabatan_pegawai & vbTab & rs!jenis_kelamin _
                & vbTab & rs!status_pegawai & vbTab & rs!status_kawin & vbTab & rs!flg_asing _
                & vbTab & rs!status_ptkp & vbTab & rs!jumlah_tanggungan & vbTab & rs!masa_perolehan_1 _
                & vbTab & rs!masa_perolehan_2 & vbTab & rs!A1 & vbTab & rs!flg_a2 _
                & vbTab & rs!A2 & vbTab & rs!A3 & vbTab & rs!a4 & vbTab & rs!a5 _
                & vbTab & rs!a6 & vbTab & rs!a7 & vbTab & rs!a8 & vbTab & rs!a9 _
                & vbTab & rs!a10 & vbTab & rs!a11 & vbTab & rs!a12 & vbTab & rs!a13 _
                & vbTab & rs!a14 & vbTab & rs!a15 & vbTab & rs!a16 & vbTab & rs!a17 _
                & vbTab & rs!a18 & vbTab & rs!a19 & vbTab & rs!a20 & vbTab & rs!a21 _
                & vbTab & rs!a22 & vbTab & rs!a22a & vbTab & rs!a22b & vbTab & rs!a23 _
                & vbTab & rs!a24 & vbTab & rs!flg_a24 & vbTab & rs!bln_a24
            rs.MoveNext
        Wend
    End If
    
    rs.Close
    LynxGrid1.Redraw = True
End Sub

Private Function checkFile(nmFile As String) As Boolean
    If Dir$(nmFile) <> "" Then
        checkFile = True
    Else
        checkFile = False
    End If
End Function
