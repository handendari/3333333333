VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_pensiun 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compensation"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtkdkar 
      Height          =   315
      Left            =   4620
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   210
      Visible         =   0   'False
      Width           =   1575
   End
   Begin prj_genting.LynxGrid LynxGrid2 
      Height          =   4815
      Left            =   1320
      TabIndex        =   40
      Top             =   1080
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8493
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
   Begin prj_genting.EnterNum txtgaji_bersih 
      Height          =   285
      Left            =   1320
      TabIndex        =   36
      Top             =   2220
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
   Begin prj_genting.vbButton vbButton2 
      Height          =   285
      Left            =   2850
      TabIndex        =   35
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
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   4530
      TabIndex        =   31
      Top             =   4800
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox txtstart_working 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1320
      TabIndex        =   25
      Top             =   1800
      Width           =   2205
   End
   Begin VB.TextBox txtdept_code 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   7020
      TabIndex        =   23
      Top             =   1110
      Width           =   885
   End
   Begin VB.TextBox txtsite_code 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   1320
      TabIndex        =   22
      Top             =   1110
      Width           =   885
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1320
      TabIndex        =   21
      Top             =   180
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   154206211
      CurrentDate     =   40823
   End
   Begin VB.TextBox txtnmDept 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   7920
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   20
      Top             =   1110
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   150
      TabIndex        =   19
      Top             =   4530
      Width           =   10695
   End
   Begin VB.TextBox txtkdtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   6540
      TabIndex        =   18
      Top             =   1470
      Width           =   885
   End
   Begin VB.TextBox txtkddiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   1470
      Width           =   735
   End
   Begin VB.TextBox txtnik 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   780
      Width           =   1515
   End
   Begin VB.TextBox txtnmkar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   3210
      TabIndex        =   12
      Top             =   780
      Width           =   4005
   End
   Begin VB.TextBox txtdivision 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2070
      TabIndex        =   11
      Top             =   1470
      Width           =   2535
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   7440
      TabIndex        =   10
      Top             =   1470
      Width           =   3195
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   41
      Top             =   2580
      Width           =   9015
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   8
      Top             =   480
      Width           =   2025
   End
   Begin VB.TextBox txtentry 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   6
      Top             =   150
      Width           =   2025
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2220
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1110
      Width           =   3525
   End
   Begin VB.PictureBox txt_jml_15_persen 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   345
      Left            =   5490
      ScaleHeight     =   285
      ScaleWidth      =   2835
      TabIndex        =   30
      Top             =   3090
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox txt_jml_pensiun 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   345
      Left            =   5460
      ScaleHeight     =   285
      ScaleWidth      =   2835
      TabIndex        =   32
      Top             =   3660
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox txt_15_persen 
      Height          =   345
      Left            =   4440
      ScaleHeight     =   285
      ScaleWidth      =   645
      TabIndex        =   1
      Top             =   3090
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   37
      Top             =   0
      Width           =   0
   End
   Begin prj_genting.EnterNum txt_kali 
      Height          =   285
      Left            =   4530
      TabIndex        =   38
      Top             =   2220
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
   Begin prj_genting.EnterNum txt_jml_pesangon 
      Height          =   285
      Left            =   5550
      TabIndex        =   39
      Top             =   2220
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
   Begin prj_genting.vbButton vbButton1 
      Height          =   585
      Left            =   900
      TabIndex        =   42
      Top             =   5160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1032
      BTYPE           =   14
      TX              =   "SAVE"
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
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_genting.vbButton vbButton5 
      Height          =   585
      Left            =   2760
      TabIndex        =   43
      Top             =   5160
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1032
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
      MICON           =   "frm_trans_pensiun.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label17 
      Caption         =   "* yyyy-MM-dd"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2970
      TabIndex        =   44
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Label16 
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
      Left            =   5190
      TabIndex        =   34
      Top             =   3150
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label15 
      Caption         =   "Total Pension :"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "Transport Compansation :"
      Height          =   255
      Left            =   90
      TabIndex        =   29
      Top             =   3120
      Visible         =   0   'False
      Width           =   1905
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
      Left            =   5280
      TabIndex        =   28
      Top             =   2250
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
      Left            =   4260
      TabIndex        =   27
      Top             =   2250
      Width           =   225
   End
   Begin VB.Label Label8 
      Caption         =   "Pension Value :"
      Height          =   255
      Left            =   150
      TabIndex        =   26
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "Start Working :"
      Height          =   210
      Left            =   210
      TabIndex        =   24
      Top             =   1830
      Width           =   1065
   End
   Begin VB.Label Label5 
      Caption         =   "Employee :"
      Height          =   210
      Left            =   480
      TabIndex        =   16
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label6 
      Caption         =   "Division :"
      Height          =   210
      Left            =   630
      TabIndex        =   15
      Top             =   1500
      Width           =   675
   End
   Begin VB.Label Label7 
      Caption         =   "Title :"
      Height          =   210
      Left            =   6090
      TabIndex        =   14
      Top             =   1530
      Width           =   405
   End
   Begin VB.Label Label9 
      Caption         =   "Remark :"
      Height          =   210
      Left            =   630
      TabIndex        =   13
      Top             =   2640
      Width           =   645
   End
   Begin VB.Label Label11 
      Caption         =   "User Name"
      Height          =   210
      Left            =   7830
      TabIndex        =   9
      Top             =   540
      Width           =   885
   End
   Begin VB.Label Label10 
      Caption         =   "Entry Date"
      Height          =   210
      Left            =   7830
      TabIndex        =   7
      Top             =   210
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Department :"
      Height          =   195
      Left            =   6060
      TabIndex        =   5
      Top             =   1170
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Company :"
      Height          =   195
      Left            =   510
      TabIndex        =   4
      Top             =   1170
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "Date :"
      Height          =   195
      Left            =   510
      TabIndex        =   2
      Top             =   240
      Width           =   495
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
      .AddColumn "NIK", 700, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "site code", , , , , , , , , False
      .AddColumn "Site", 1300, , , , , , , , , True
      .AddColumn "Depart. code", , , , , , , , , False
      .AddColumn "Departement", 1300, , , , , , , , , True
      .AddColumn "Div. code", , , , , , , , , False
      .AddColumn "Division", 1300, , , , , , , , , True
      .AddColumn "title code", , , , , , , , , False
      .AddColumn "Title", 1300, , , , , , , , , True
      .AddColumn "start_working", , , lgDate, "dd-MMM-yyyy", , , , , False
      .AddColumn "salary", , , lgNumeric, , , , , , False
      .AddColumn "Employee Code", 3000, , , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
Dim access As String

    If pilihan = 1 Then
        LynxGrid2.Clear
        strsql = "SELECT a.nik,a.employee_name," _
                    & "a.company_code,e.company_name," _
                    & "a.department_code,d.department_name," _
                    & "a.division_code,b.division_name," _
                    & "a.title_code,c.title_name,a.start_working," _
                    & "(select main_salary from m_salary_standard where employee_code = a.employee_code order by number desc limit 1) salary, a.employee_code " _
                & "FROM m_employee a LEFT JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                & "LEFT JOIN m_title c ON a.title_code = c.title_code " _
                & "LEFT JOIN m_department d ON a.department_code = d.department_code and d.company_code = a.company_code " _
                & "JOIN m_company e ON a.company_code = e.company_code " _
                & "WHERE (a.nik LIKE '%" & txtnik.Text & "%' " _
                & "OR a.employee_name LIKE '%" & txtnik.Text & "%') " _
                & "AND a.flag_active <> 0 AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                & "ORDER BY a.company_code, a.department_code, a.division_code, a.employee_name ASC"
                
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid2.AddItem rs2!nik & vbTab & rs2!EMPLOYEE_NAME & vbTab & _
                    rs2!COMPANY_CODE & vbTab & rs2!company_name & vbTab & rs2!DEPARTMENT_CODE & vbTab & rs2!department_name & vbTab & _
                    rs2!DIVISION_CODE & vbTab & rs2!division_name & vbTab & rs2!title_code & vbTab & _
                    rs2!title_name & vbTab & rs2!start_working & vbTab & rs2!salary & vbTab & rs2!employee_code
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
                txtdept_code.Text = rs2!DEPARTMENT_CODE
                txtnmDept.Text = rs2!department_name
                txtkddiv.Text = rs2!DIVISION_CODE
                txtdivision.Text = IIf(IsNull(rs2!division_name) = True, "", rs2!division_name)
                txtkdtitle.Text = rs2!title_code
                txttitle.Text = rs2!title_name
                txtstart_working.Text = Format(rs2!start_working, "dd-MMM-yyyy")
                txtgaji_bersih.Value = rs2!salary
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
            txtdept_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 4)
            txtnmDept.Text = LynxGrid2.CellText(LynxGrid2.Row, 5)
            txtkddiv.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
            txtdivision.Text = LynxGrid2.CellText(LynxGrid2.Row, 7)
            txtkdtitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 8)
            txttitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 9)
            txtstart_working.Text = LynxGrid2.CellText(LynxGrid2.Row, 10)
            txtgaji_bersih.Value = LynxGrid2.CellValue(LynxGrid2.Row, 11)
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

    '+++++++++++++++++++APAKAH KODE KARYAWAN SUDAH BENAR++++++++++++++++++++
    strsql = "SELECT employee_code FROM m_employee WHERE employee_code = '" & txtkdkar.Text & "'"
    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs2.RecordCount = 0 Then
        MsgBox "Invalid Employee Code...!!", vbCritical
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
            MsgBox "This Employee Is Already Pension...!!", vbCritical
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
                "fn_GetExpenseothers(a.employee_code,'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "') potongan_lain," & _
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
        v_jamsostek = txtgaji_bersih.Value * 0.02
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
    
    Dim tglStart As String
    Dim tglEnd As String
    
    tglStart = Format(DTPicker1.Value, "yyyy-MM-01")
    tglEnd = Format(DTPicker1.Value, "yyyy-MM-") & getEndDay(month(DTPicker1.Value), year(DTPicker1.Value))
    
    strsql = "DELETE FROM h_salary WHERE employee_code = '" & txtkdkar.Text & "' " _
            & "AND month = '" & Mid(Format(DTPicker1.Value, "yyyy-MM-dd"), 1, 7) & "'"
    CnG.Execute (strsql)

    bulan = Mid(Format(DTPicker1.Value, "yyyy-MM-dd"), 1, 7)

    strsql = "Select a.employee_code, a.employee_name," & _
            "a.marital_status, a.number_of_children, a.sex," & _
            "a.no_jamsostek,a.npwp,a.company_code," & _
            "(SELECT pph21_type FROM m_salary_standard WHERE employee_code = a.employee_code ORDER BY number DESC LIMIT 1) pph21_type, " & _
            "(SELECT jstk_type FROM m_salary_standard WHERE employee_code = a.employee_code ORDER BY number DESC LIMIT 1) jstk_type " & _
        "FROM m_employee a " & _
        "WHERE employee_code = '" & txtkdkar.Text & "'"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

    Call clsinsert.CalcSuFormula(bulan, rs!employee_code, rs!no_jamsostek, rs!npwp, rs!COMPANY_CODE, rs!pph21_type, rs!jstk_type, tglStart, tglEnd)

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
