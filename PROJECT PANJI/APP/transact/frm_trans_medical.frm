VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_medical 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CUTI SAKIT BERBAYAR"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prj_panji.LynxGrid LynxGrid2 
      Height          =   3165
      Left            =   1470
      TabIndex        =   25
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
   Begin VB.Frame frmStatus 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   1470
      TabIndex        =   29
      Top             =   2160
      Width           =   3195
      Begin MSComCtl2.DTPicker DTPicker_sakit 
         Height          =   315
         Left            =   1590
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   32505859
         CurrentDate     =   41288
      End
      Begin VB.OptionButton optSembuh 
         Caption         =   "SEMBUH"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   1305
      End
      Begin VB.OptionButton optSakit 
         Caption         =   "SAKIT"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Value           =   -1  'True
         Width           =   915
      End
      Begin MSComCtl2.DTPicker DTPicker_sembuh 
         Height          =   315
         Left            =   1590
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   32505859
         CurrentDate     =   41288
      End
   End
   Begin VB.TextBox txtkdkar 
      Height          =   285
      Left            =   4560
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   300
      Visible         =   0   'False
      Width           =   2205
   End
   Begin prj_panji.vbButton vbButton2 
      Height          =   285
      Left            =   3000
      TabIndex        =   23
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
      MICON           =   "frm_trans_medical.frx":0000
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
      TabIndex        =   22
      Top             =   2130
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.TextBox txtsite_code 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   1470
      TabIndex        =   20
      Top             =   1110
      Width           =   885
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1470
      TabIndex        =   19
      Top             =   180
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   32505859
      CurrentDate     =   40823
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   150
      TabIndex        =   18
      Top             =   4020
      Width           =   10695
   End
   Begin VB.TextBox txtkdtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1470
      TabIndex        =   17
      Top             =   1800
      Width           =   885
   End
   Begin VB.TextBox txtkddiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1470
      TabIndex        =   16
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
      TabIndex        =   11
      Top             =   780
      Width           =   4005
   End
   Begin VB.TextBox txtdivision 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2370
      TabIndex        =   10
      Top             =   1470
      Width           =   3525
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2370
      TabIndex        =   9
      Top             =   1800
      Width           =   3525
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1470
      MaxLength       =   50
      TabIndex        =   5
      Top             =   3150
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
      TabIndex        =   7
      Top             =   1110
      Width           =   3525
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   24
      Top             =   0
      Width           =   0
   End
   Begin prj_panji.vbButton vbButton1 
      Height          =   705
      Left            =   690
      TabIndex        =   27
      Top             =   4530
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
      MICON           =   "frm_trans_medical.frx":001C
      PICN            =   "frm_trans_medical.frx":0038
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
      TabIndex        =   28
      Top             =   4530
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
      MICON           =   "frm_trans_medical.frx":10CA
      PICN            =   "frm_trans_medical.frx":10E6
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
      Alignment       =   1  'Right Justify
      Caption         =   "TGL MASUK"
      Height          =   210
      Left            =   90
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "KARYAWAN"
      Height          =   210
      Left            =   150
      TabIndex        =   15
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "DIVISI"
      Height          =   210
      Left            =   480
      TabIndex        =   14
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "JABATAN"
      Height          =   210
      Left            =   150
      TabIndex        =   13
      Top             =   1860
      Width           =   1245
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "KETERANGAN"
      Height          =   210
      Left            =   180
      TabIndex        =   12
      Top             =   3210
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   1170
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TANGGAL"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frm_trans_medical"
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
      .AddColumn "Employee Code", 3000, , , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
Dim access As String

    If pilihan = 1 Then
        LynxGrid2.Clear
        If LOGIN_LEVEL = 100 Then
            strsql = "SELECT a.nik,a.employee_name," _
                        & "a.company_code,e.company_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.start_working,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE (a.nik LIKE '%" & txtnik.Text & "%' " _
                    & "OR a.employee_name LIKE '%" & txtnik.Text & "%') " _
                    & "AND a.flag_active <> 0"
        Else
            strsql = "SELECT a.nik,a.employee_name," _
                        & "a.company_code,e.company_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.start_working,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE (a.nik LIKE '%" & txtnik.Text & "%' " _
                    & "OR a.employee_name LIKE '%" & txtnik.Text & "%') " _
                    & "AND a.flag_active <> 0 AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                    & "ORDER BY a.company_code, a.department_code, a.division_code, a.employee_name ASC"

        End If
        
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid2.AddItem rs2!nik & vbTab & rs2!EMPLOYEE_NAME & vbTab & _
                    rs2!COMPANY_CODE & vbTab & rs2!company_name & vbTab & _
                    rs2!division_code & vbTab & rs2!division_name & vbTab & rs2!title_code & vbTab & _
                    rs2!title_name & vbTab & rs2!start_working & vbTab & rs2!employee_code
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
                txtket.SetFocus
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
            txtkdkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 9)
            txtket.SetFocus
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
    optSakit_Click
    
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

Private Sub optSakit_Click()
    If optSakit.Value = True Then
        DTPicker_sakit.Visible = True
        DTPicker_sakit.Value = Now
    Else
        DTPicker_sakit.Visible = False
    End If
End Sub

Private Sub optSembuh_Click()
    If optSembuh.Value = True Then
        DTPicker_sembuh.Visible = True
        DTPicker_sembuh.Value = Now
    Else
        DTPicker_sembuh.Visible = False
    End If
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
Dim i As Integer
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
        strsql = "SELECT tgltrans FROM t_medical WHERE employee_code = '" & txtkdkar.Text & "'"
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            MsgBox "Karyawan Sudah Ada...", vbCritical
            rs2.Close
            Exit Sub
        End If
        rs2.Close
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    i = MsgBox("Data Kehadiran Setelah Tanggal Input Akan Dihapus" & Chr(13) & _
        "Apakah Akan Melanjutkan Proses?", vbYesNo, headerMSG)
    If i = vbYes Then
        strsql = "DELETE FROM h_attendance " & _
                "WHERE employee_code = '" & txtkdkar.Text & "' " & _
                "AND DATE(att_date) >= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
        CnG.Execute strsql
        MsgBox "Data Kehadiran Berhasil Dihapus...", vbInformation, headerMSG
    Else
        MsgBox "Data Kehadiran Batal Dihapus...", vbInformation, headerMSG
        Unload Me
        frm_List_medical.isiGridAbsen
        Exit Sub
    End If
    
    If editTrans = False Then
        strsql = "INSERT INTO t_medical " & _
                "(tgltrans,employee_code,ket,userinput,tglinput,flag_sakit,tgl_sakit) " & _
                "VALUES " & _
                "('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & txtkdkar.Text & "'," & _
                "'" & txtket.Text & "','" & LOGIN_CODE & "',now(),'" & IIf(optSakit.Value, 1, 0) & "','" & Format(DTPicker_sakit.Value, "yyyy-MM-dd") & "')"
    Else
        strsql = "UPDATE t_medical SET " & _
                "tgltrans = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',employee_code = '" & txtkdkar.Text & "'," & _
                "ket = '" & txtket.Text & "',useredit = '" & LOGIN_CODE & "',tgledit = now(), " & _
                "flag_sembuh = '" & IIf(optSembuh.Value, 1, 0) & "',tgl_sembuh = '" & Format(DTPicker_sembuh.Value, "yyyy-MM-dd") & "' " & _
                "WHERE employee_code = '" & txtkdkar.Text & "'"
    End If
    
    CnG.Execute strsql
    
    txtkdkar.Text = ""
    txtnmkar.Text = ""
    txtkddiv.Text = ""
    txtdivision.Text = ""
    txtkdtitle.Text = ""
    txttitle.Text = ""
    txtstart_working.Text = ""
    txtsite_code.Text = ""
    txt_company_name.Text = ""
    txtket.Text = ""
    txtnik.SetFocus
    
    MsgBox "Data Berhasil Disimpan...", vbInformation, headerMSG
    
    Unload Me
    frm_List_medical.isiGridAbsen
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
