VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_pensiun 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KOMPENSASI"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtkdkar 
      Height          =   285
      Left            =   4560
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   300
      Visible         =   0   'False
      Width           =   345
   End
   Begin prj_panji.LynxGrid LynxGrid2 
      Height          =   3165
      Left            =   1470
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   5583
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
   Begin prj_panji.EnterNum txtgaji_bersih 
      Height          =   285
      Left            =   1470
      TabIndex        =   1
      Top             =   2520
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prj_panji.vbButton vbButton2 
      Height          =   285
      Left            =   3000
      TabIndex        =   24
      Top             =   780
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
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
      MICON           =   "frm_trans_pensiun.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtstart_working 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1470
      TabIndex        =   20
      Top             =   2130
      Width           =   2205
   End
   Begin VB.TextBox txtsite_code 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   1470
      TabIndex        =   18
      Top             =   1110
      Width           =   885
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1470
      TabIndex        =   17
      Top             =   180
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   95879171
      CurrentDate     =   40823
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   150
      TabIndex        =   16
      Top             =   3660
      Width           =   10695
   End
   Begin VB.TextBox txtkdtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1470
      TabIndex        =   15
      Top             =   1800
      Width           =   885
   End
   Begin VB.TextBox txtkddiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1470
      TabIndex        =   14
      Top             =   1470
      Width           =   885
   End
   Begin VB.TextBox txtnik 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1470
      TabIndex        =   0
      Top             =   780
      Width           =   1515
   End
   Begin VB.TextBox txtnmkar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Top             =   780
      Width           =   4005
   End
   Begin VB.TextBox txtdivision 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2370
      TabIndex        =   8
      Top             =   1470
      Width           =   3525
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2370
      TabIndex        =   7
      Top             =   1800
      Width           =   3525
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1470
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2880
      Width           =   9015
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2370
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1110
      Width           =   3525
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   25
      Top             =   0
      Width           =   0
   End
   Begin prj_panji.EnterNum txt_kali 
      Height          =   285
      Left            =   4710
      TabIndex        =   2
      Top             =   2520
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prj_panji.EnterNum txt_jml_pesangon 
      Height          =   285
      Left            =   5700
      TabIndex        =   26
      Top             =   2520
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prj_panji.vbButton vbButton1 
      Height          =   705
      Left            =   690
      TabIndex        =   29
      Top             =   4320
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1244
      BTYPE           =   14
      TX              =   "&Save"
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
      MICON           =   "frm_trans_pensiun.frx":001C
      PICN            =   "frm_trans_pensiun.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.vbButton vbButton5 
      Height          =   705
      Left            =   2340
      TabIndex        =   30
      Top             =   4320
      Width           =   1365
      _ExtentX        =   2408
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
      MICON           =   "frm_trans_pensiun.frx":10CA
      PICN            =   "frm_trans_pensiun.frx":10E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5430
      TabIndex        =   23
      Top             =   2550
      Width           =   225
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4410
      TabIndex        =   22
      Top             =   2550
      Width           =   225
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "KOMPENSASI"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   2550
      Width           =   1395
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "TGL MASUK"
      Height          =   210
      Left            =   90
      TabIndex        =   19
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "KARYAWAN"
      Height          =   210
      Left            =   150
      TabIndex        =   13
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "DIVISI"
      Height          =   210
      Left            =   480
      TabIndex        =   12
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "JABATAN"
      Height          =   210
      Left            =   150
      TabIndex        =   11
      Top             =   1860
      Width           =   1245
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Remark :"
      Height          =   210
      Left            =   540
      TabIndex        =   10
      Top             =   2940
      Width           =   885
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   1170
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TANGGAL"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frm_trans_pensiun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs2 As New ADODB.Recordset
Dim strsql As String
Public editTrans As Boolean

Private Sub createKar()
   With LynxGrid2
      .AddColumn "KODE KARY.", 700, lgAlignCenterCenter, , , , , , , True
      .AddColumn "NAMA KARY.", 3000, , , , , , , , , True
      .AddColumn "site code", , , , , , , , , False
      .AddColumn "PERUSAHAAN", 1300, , , , , , , , , True
      .AddColumn "Div. code", , , , , , , , , False
      .AddColumn "DIVISI", 1300, , , , , , , , , True
      .AddColumn "title code", , , , , , , , , False
      .AddColumn "JABATAN", 1300, , , , , , , , , True
      .AddColumn "TGL MASUK", , , lgDate, "dd-MMM-yyyy", , , , , False
      .AddColumn "GAJI POKOK", , , lgNumeric, , , , , , False
      .AddColumn "FLAG BASIC", , , lgNumeric, , , , , , False
      .AddColumn "JHK VALUE", , , , , , , , , True
      .AddColumn "Employee Code", 3000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
Dim access As String
Dim vFlagBasic As Integer
Dim vJHK As Double

    If pilihan = 1 Then
        LynxGrid2.Clear
        If LOGIN_LEVEL = 100 Then
            strsql = "SELECT a.nik,a.employee_name," _
                        & "a.company_code,e.company_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.start_working," _
                        & "(SELECT f.basic_salary FROM m_salary_standard f WHERE f.employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) salary," _
                        & "(SELECT f.flag_basic FROM m_salary_standard f WHERE f.employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) flag_basic," _
                        & "d.jhk_value,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "JOIN m_pref_jhk d ON a.division_code = d.division_code " _
                    & "WHERE (a.nik LIKE '%" & txtnik.Text & "%' " _
                    & "OR a.employee_name LIKE '%" & txtnik.Text & "%') " _
                    & "AND a.flag_active <> 0"
        Else
            strsql = "SELECT a.nik,a.employee_name," _
                        & "a.company_code,e.company_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.start_working," _
                        & "(SELECT f.basic_salary FROM m_salary_standard f WHERE f.employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) salary," _
                        & "(SELECT f.flag_basic FROM m_salary_standard f WHERE f.employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) flag_basic," _
                        & "d.jhk_value,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "JOIN m_pref_jhk d ON a.division_code = d.division_code " _
                    & "WHERE (a.nik LIKE '%" & txtnik.Text & "%' " _
                    & "OR a.employee_name LIKE '%" & txtnik.Text & "%') " _
                    & "AND a.flag_active <> 0 AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                    & "ORDER BY a.company_code, a.division_code, a.employee_name ASC"

        End If
        
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid2.AddItem rs2!nik & vbTab & rs2!EMPLOYEE_NAME & vbTab & _
                    rs2!COMPANY_CODE & vbTab & rs2!company_name & vbTab & _
                    rs2!division_code & vbTab & rs2!division_name & vbTab & rs2!title_code & vbTab & _
                    rs2!title_name & vbTab & rs2!start_working & vbTab & rs2!salary & vbTab & _
                    rs2!flag_basic & vbTab & rs2!jhk_value & vbTab & rs2!employee_code
                rs2.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs2.RecordCount = 1 Then
                rs2.MoveFirst
                txtkdkar.Text = rs2!employee_code
                txtnik.Text = rs2!nik
                txtnmkar.Text = rs2!EMPLOYEE_NAME
                txtsite_code.Text = rs2!COMPANY_CODE
                txt_company_name.Text = rs2!company_name
                txtkddiv.Text = rs2!division_code
                txtdivision.Text = IIf(IsNull(rs2!division_name) = True, "", rs2!division_name)
                txtkdtitle.Text = rs2!title_code
                txttitle.Text = rs2!title_name
                txtstart_working.Text = Format(rs2!start_working, "dd-MMM-yyyy")
                
                vFlagBasic = IIf(IsNull(rs2!flag_basic), 0, rs2!flag_basic)
                vJHK = rs2!jhk_value
                
                If vFlagBasic = 0 Then
                    txtgaji_bersih.Value = IIf(IsNull(rs2!salary), 0, rs2!salary) * vJHK
                Else
                    txtgaji_bersih.Value = IIf(IsNull(rs2!salary), 0, rs2!salary)
                End If
                
                txtgaji_bersih.SetFocus
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else
            
        End If
        rs2.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txtnik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txtnmkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txtsite_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
            txt_company_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 3)
            txtkddiv.Text = LynxGrid2.CellText(LynxGrid2.Row, 4)
            txtdivision.Text = LynxGrid2.CellText(LynxGrid2.Row, 5)
            txtkdtitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
            txttitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 7)
            txtstart_working.Text = LynxGrid2.CellText(LynxGrid2.Row, 8)
            
            vFlagBasic = LynxGrid2.CellValue(LynxGrid2.Row, 10)
            vJHK = LynxGrid2.CellValue(LynxGrid2.Row, 11)
                
            If vFlagBasic = 0 Then
                txtgaji_bersih.Value = LynxGrid2.CellValue(LynxGrid2.Row, 9) * vJHK
            Else
                txtgaji_bersih.Value = LynxGrid2.CellValue(LynxGrid2.Row, 9)
            End If
            
            txtkdkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 12)
            txtgaji_bersih.SetFocus
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        vbButton1_Click
    End If
End Sub

Private Sub Form_Load()
Dim rsshift As New ADODB.Recordset
    createKar
    
    DTPicker1.Value = Now
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

'Private Sub txt_15_persen_Validate(Cancel As Boolean)
'    txt_jml_15_persen.Value = (txt_15_persen.Value / 100) * txt_jml_pesangon.Value
'    txt_jml_pensiun.Value = txt_jml_pesangon.Value + txt_jml_15_persen.Value
'End Sub

Private Sub txt_kali_Validate(Cancel As Boolean)
    txt_jml_pesangon.Value = txtgaji_bersih.Value * txt_kali.Value
'    txt_jml_15_persen.Value = (txt_15_persen.Value / 100) * txt_jml_pesangon.Value
'    txt_jml_pensiun.Value = txt_jml_pesangon.Value + txt_jml_15_persen.Value
End Sub

Private Sub txtgaji_bersih_Validate(Cancel As Boolean)
    txt_jml_pesangon.Value = txtgaji_bersih.Value * txt_kali.Value
'    txt_jml_15_persen.Value = (txt_15_persen.Value / 100) * txt_jml_pesangon.Value
'    txt_jml_pensiun.Value = txt_jml_pesangon.Value + txt_jml_15_persen.Value
End Sub

Private Sub txtkdkar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys (vbTab)
    End If
End Sub

Private Sub vbButton1_Click()
Dim rs As New Recordset
Dim strsql As String
Dim vJamsostek As Double
Dim vTipeJstk As String

    '+++++++++++++++++++APAKAH KODE KARYAWAN SUDAH BENAR++++++++++++++++++++
    strsql = "SELECT employee_code FROM m_employee WHERE employee_code = '" & txtkdkar.Text & "'"
    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs2.RecordCount = 0 Then
        MsgBox "Kode Karyawan Tidak Valid...", vbCritical
        rs2.Close
        Exit Sub
    End If
    rs2.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++APAKAH KARYAWAN SUDAH PERNAH DI INPUT++++++++++++++++++++
    If editTrans = False Then
        strsql = "SELECT tgltrans FROM t_pensiun WHERE employee_code = '" & txtkdkar.Text & "'"
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            MsgBox "Karyawan Sudah Menerima Kompensasi...", vbCritical
            rs2.Close
            Exit Sub
        End If
        rs2.Close
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++++++++++++ Proses Kalkulasi Deduction ++++++++++++++++++++
    Dim rsproses As New ADODB.Recordset
    Dim v_pph_0, v_pph_5, v_pph_15, v_pph_25 As Double
    Dim v_pph21_pension As Double
    Dim v_loan, v_jamsostek, v_absen, v_others, v_absen_others As Double
    Dim v_tot_income, v_tot_deduction, v_total As Double
    
    strsql = "SELECT " & _
                "fn_jmlHariKerja(a.employee_code,'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "') jmlHariMasuk, " & _
                "fn_jmlHariAbsent(a.employee_code,'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "') jmlHariAbsent," & _
                "fn_GetExpenseOthers(a.employee_code,'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "') potongan_lain," & _
                "SUM(d.installment_amount) loan " & _
            "FROM m_employee a " & _
            "JOIN td_loan d ON a.employee_code = d.employee_code " & _
            "WHERE a.employee_code = '" & txtkdkar.Text & "' AND a.company_code = '" & txtsite_code.Text & "' " & _
            "AND d.flag_paid = 0"
    rsproses.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
    If rsproses.RecordCount > 0 Then
        Dim rspph As New ADODB.Recordset
        
'        Dim rsincome As New ADODB.Recordset
'        Dim v_masa_kerja As Double
'        Dim v_wisdom As Double
'        Dim v_leave As Double
'        Dim v_leave_value As Double
'
'        strsql = "SELECT PERIOD_DIFF(YEAR(NOW()),YEAR(a.start_working)) AS masa_kerja, " _
'            & "(SELECT over_leave FROM t_leave_periode WHERE employee_code = a.employee_code ORDER BY start_periode, end_periode DESC LIMIT 1) * -1 AS sisa_cuti " _
'            & "FROM m_employee a WHERE a.employee_code = '" & txtkdkar.Text & "'"
'        rsincome.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'        If rsincome.RecordCount > 0 Then
'            v_masa_kerja = rsincome!masa_kerja
'            v_wisdom = 2 * txtgaji_bersih.Value * v_masa_kerja
'            v_leave = rsincome!sisa_cuti
'            v_leave_value = (txtgaji_bersih.Value / 25) * v_leave
'        End If
'        rsincome.Close
        
        v_loan = IIf(IsNull(rsproses!loan), 0, rsproses!loan)
        
        SQL = "SELECT jstk_type FROM m_salary_standard WHERE employee_code = '" & txtkdkar.Text & "' AND date(salary_date) <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
                "ORDER BY salary_date DESC LIMIT 1"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            vTipeJstk = rs!jstk_type
        End If
        rs.Close
        
        SQL = "SELECT tk FROM m_jamsostek_detail WHERE jamsostek_code = '" & vTipeJstk & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            vJamsostek = IIf(IsNull(rs!tk), 0, rs!tk)
        Else
            vJamsostek = 0
        End If
        rs.Close
        
        v_jamsostek = txtgaji_bersih.Value * (vJamsostek / 100)
        v_others = IIf(IsNull(rsproses!potongan_lain), 0, rsproses!potongan_lain)
'        v_absen = (rsproses!jmlHariAbsent / 25) * txtgaji_bersih.Value
'        v_absen_others = (v_others + v_absen)
        
'        v_tot_income = (txt_jml_pesangon.Value + txt_jml_15_persen.Value + v_wisdom + v_leave_value)
        
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 1 AND pph21_code = 'CPST'"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If txt_jml_pesangon.Value > rspph!pph21_upper Then '50000000
                v_pph_0 = (rspph!pph21_percentage / 100) * rspph!pph21_upper '50000000
            Else
                v_pph_0 = (rspph!pph21_percentage / 100) * txt_jml_pesangon.Value
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 2 AND pph21_code = 'CPST'"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If txt_jml_pesangon.Value <= rspph!pph21_under Then '50000000
                v_pph_5 = 0
            ElseIf v_tot_income < rspph!pph21_upper Then '100000000
                v_pph_5 = (rspph!pph21_percentage / 100) * (txt_jml_pesangon.Value - rspph!pph21_under) '50000000
            Else
'                v_pph_5 = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '50000000
                v_pph_5 = (rspph!pph21_percentage / 100) * (txt_jml_pesangon.Value - rspph!pph21_under) '50000000
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 3 AND pph21_code = 'CPST'"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If txt_jml_pesangon.Value <= rspph!pph21_under Then '500000000
                v_pph_15 = 0
            ElseIf txt_jml_pesangon.Value < rspph!pph21_upper Then  '500000000
                v_pph_15 = (rspph!pph21_percentage / 100) * (txt_jml_pesangon.Value - rspph!pph21_under) '100000000
            Else
                v_pph_15 = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '100000000
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 4 AND pph21_code = 'CPST'"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If txt_jml_pesangon.Value <= rspph!pph21_under Then '500000000
                v_pph_25 = 0
            Else
                If txt_jml_pesangon.Value > rspph!pph21_under Then '500000000
                    v_pph_25 = (rspph!pph21_percentage / 100) * (v_tot_income - rspph!pph21_under) '500000000
                Else
                    v_pph_25 = 0
                End If
            End If
        End If
        rspph.Close
        
        v_pph21_pension = (v_pph_0 + v_pph_5 + v_pph_15 + v_pph_25)
        
        v_tot_deduction = (v_pph21_pension + v_loan + v_jamsostek)
        v_total = (txt_jml_pesangon.Value - v_tot_deduction)
    End If
    rsproses.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    If editTrans = False Then
        strsql = "INSERT INTO t_pensiun " & _
                "(tgltrans,employee_code,basic_salary,kali_gaji,uang_pensiun,ket,userinput,tglinput, " & _
                "pph21,loan,jamsostek,total_income,total_deduction,total) " & _
                "VALUES " & _
                "('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & txtkdkar.Text & "'," & _
                "'" & txtgaji_bersih.Value & "','" & txt_kali.Value & "','" & txt_jml_pesangon.Value & "'," & _
                "'" & txtket.Text & "','" & LOGIN_CODE & "',now(), " & _
                "'" & v_pph21_pension & "','" & v_loan & "','" & v_jamsostek & "'," & _
                "'" & txt_jml_pesangon.Value & "','" & v_tot_deduction & "','" & v_total & "')"
    Else
        strsql = "UPDATE t_pensiun SET " & _
                "tgltrans = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',employee_code = '" & txtkdkar.Text & "'," & _
                "basic_salary = '" & txtgaji_bersih.Value & "',kali_gaji = '" & txt_kali.Value & "'," & _
                "uang_pensiun = '" & txt_jml_pesangon.Value & "', " & _
                "ket = '" & txtket.Text & "',useredit = '" & LOGIN_CODE & "',tgledit = now(), " & _
                "pph21 = '" & v_pph21_pension & "',loan = '" & v_loan & "',jamsostek = '" & v_jamsostek & "'," & _
                "total_income = '" & txt_jml_pesangon.Value & "'," & _
                "total_deduction = '" & v_tot_deduction & "',total = '" & v_total & "' " & _
                "WHERE employee_code = '" & txtkdkar.Text & "'"
    End If
    
    CnG.Execute strsql
    
'    Dim clsinsert As New clsInsert_h_salary
'    Dim bulan As String
'
'    strsql = "DELETE FROM h_salary WHERE employee_code = '" & txtkdkar.Text & "' " _
'            & "AND month = '" & Mid(Format(DTPicker1.Value, "yyyy-MM-dd"), 1, 7) & "'"
'    CnG.Execute (strsql)
'
'    bulan = Mid(Format(DTPicker1.Value, "yyyy-MM-dd"), 1, 7)
'
'    strsql = "Select employee_code, employee_name," & _
'            "marital_status, number_of_children, sex," & _
'            "end_working, start_mc, CAST(IFNULL(end_mc,LAST_DAY('" & tgl & "')) as DATE) end_mc, flag_active, " & _
'            "CONCAT(DATE_FORMAT(LAST_DAY(NOW()),'%Y-%m-'),'01') tgl_awal " & _
'        "FROM m_employee " & _
'        "WHERE employee_code = '" & txtkdkar.Text & "'"
'    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'    Call clsinsert.Insert_h_salary(rs!employee_code, rs!sex, bulan, Format(rs!tgl_awal, "yyyy-MM-dd"), _
'                Format(DTPicker1.Value, "yyyy-MM-dd"), rs!marital_status, IIf(IsNull(rs!number_of_children), 0, rs!number_of_children), _
'                IIf(IsNull(Format(rs!start_mc, "yyyyMM")), "0", Format(rs!start_mc, "yyyyMM")), _
'                rs!flag_active, txtsite_code.Text)
'
'    rs.Close

    Dim clsinsert As New clsCalcSUFormula
    Dim bulan As String

    strsql = "DELETE FROM h_salary WHERE employee_code = '" & txtkdkar.Text & "' " _
            & "AND month = '" & Mid(Format(DTPicker1.Value, "yyyy-MM-dd"), 1, 7) & "'"
    CnG.Execute (strsql)

    bulan = Mid(Format(DTPicker1.Value, "yyyy-MM-dd"), 1, 7)

    strsql = "Select a.employee_code, a.employee_name," & _
            "a.marital_status, a.no_of_children, a.sex," & _
            "a.no_jamsostek,a.npwp,a.company_code," & _
            "(SELECT f.pph21_type FROM m_salary_standard f WHERE f.employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) pph21_type, " & _
            "(SELECT f.jstk_type FROM m_salary_standard f WHERE f.employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) jstk_type " & _
        "FROM m_employee a " & _
        "WHERE employee_code = '" & txtkdkar.Text & "'"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

    Call clsinsert.CalcSuFormula(bulan, rs!employee_code, rs!no_jamsostek, IIf(IsNull(rs!npwp), "", rs!npwp), rs!COMPANY_CODE, IIf(IsNull(rs!pph21_type), 0, rs!pph21_type), IIf(IsNull(rs!jstk_type), "STD", rs!jstk_type))

    rs.Close
    
    strsql = "UPDATE m_employee SET flag_active = 0 WHERE employee_code = '" & txtkdkar.Text & "'"
    CnG.Execute strsql
    
    txtkdkar.Text = ""
    txtnmkar.Text = ""
    txtkddiv.Text = ""
    txtdivision.Text = ""
    txtkdtitle.Text = ""
    txttitle.Text = ""
    txtgaji_bersih.Value = 0
    txt_jml_pesangon.Value = 0
'    txt_jml_pensiun.Value = 0
'    txt_jml_15_persen.Value = 0
    txt_kali.Value = 0
    txtstart_working.Text = ""
    txtsite_code.Text = ""
    txt_company_name.Text = ""
    txtket.Text = ""
    txtnik.SetFocus
    
    MsgBox "Save Succesfully!", vbInformation, "Success!"
    
    Unload Me
    frm_List_pensiun.isiGridAbsen
End Sub

Private Sub vbbutton2_Click()
    isiGridKar (1)
End Sub

Private Sub vbButton5_Click()
    Unload Me
End Sub

Private Sub txtnik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub
