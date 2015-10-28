VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_IncomeLain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OTHER EMPLOYEE  INCOME "
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10095
   Icon            =   "frm_trans_IncomeLain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   4365
      Left            =   2640
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7699
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
   Begin VB.TextBox txt_dep_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4470
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   18
      Top             =   1950
      Width           =   3855
   End
   Begin VB.TextBox txt_department_code 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   26
      Top             =   1950
      Width           =   1785
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frm_trans_IncomeLain.frx":058A
      Left            =   2640
      List            =   "frm_trans_IncomeLain.frx":0594
      TabIndex        =   24
      Text            =   "INCOME"
      Top             =   2940
      Width           =   1785
   End
   Begin prj_tpc.vbButton vbButton2 
      Height          =   285
      Left            =   4110
      TabIndex        =   21
      Top             =   1620
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
      MICON           =   "frm_trans_IncomeLain.frx":05A9
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
      Height          =   30
      Left            =   360
      TabIndex        =   16
      Top             =   5220
      Width           =   10695
   End
   Begin VB.TextBox txtkdtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2640
      TabIndex        =   15
      Top             =   2610
      Width           =   1785
   End
   Begin VB.TextBox txtkddiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2640
      TabIndex        =   14
      Top             =   2280
      Width           =   1785
   End
   Begin VB.TextBox txtnik 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Top             =   1620
      Width           =   1425
   End
   Begin VB.TextBox txtnmkar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   4470
      TabIndex        =   9
      Top             =   1620
      Width           =   3855
   End
   Begin VB.TextBox txtdivision 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   4470
      TabIndex        =   8
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   4470
      TabIndex        =   7
      Top             =   2610
      Width           =   3855
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   20
      Top             =   3690
      Width           =   5685
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4470
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1260
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Top             =   870
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   92340227
      CurrentDate     =   40794
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   2640
      OleObjectBlob   =   "frm_trans_IncomeLain.frx":05C5
      TabIndex        =   4
      Top             =   1260
      Width           =   1785
   End
   Begin VB.TextBox txtjumlah 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   19
      Top             =   3300
      Width           =   1785
   End
   Begin VB.TextBox txtkdkar 
      Height          =   285
      Left            =   8130
      TabIndex        =   22
      Top             =   1620
      Visible         =   0   'False
      Width           =   345
   End
   Begin prj_tpc.vbButton cmdSave 
      Height          =   705
      Left            =   600
      TabIndex        =   28
      Top             =   5670
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
      MICON           =   "frm_trans_IncomeLain.frx":252B
      PICN            =   "frm_trans_IncomeLain.frx":2547
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   2250
      TabIndex        =   29
      Top             =   5670
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
      MICON           =   "frm_trans_IncomeLain.frx":35D9
      PICN            =   "frm_trans_IncomeLain.frx":35F5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frm_trans_IncomeLain.frx":4687
      Left            =   4470
      List            =   "frm_trans_IncomeLain.frx":4691
      TabIndex        =   30
      Top             =   2940
      Width           =   1785
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT OTHER INCOME / EXPENSE"
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
      TabIndex        =   27
      Top             =   150
      Width           =   5925
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "TYPE :"
      Height          =   210
      Left            =   1590
      TabIndex        =   25
      Top             =   3000
      Width           =   945
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "VALUE :"
      Height          =   195
      Left            =   1560
      TabIndex        =   17
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "EMPLOYEE :"
      Height          =   210
      Left            =   1560
      TabIndex        =   13
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "SECTION :"
      Height          =   210
      Left            =   1740
      TabIndex        =   12
      Top             =   2370
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "JOB TITLE :"
      Height          =   210
      Left            =   1650
      TabIndex        =   11
      Top             =   2670
      Width           =   885
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "DESCRIPTION :"
      Height          =   210
      Left            =   1110
      TabIndex        =   10
      Top             =   3720
      Width           =   1425
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DEPARTMENT :"
      Height          =   195
      Left            =   1320
      TabIndex        =   6
      Top             =   2010
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "COMPANY :"
      Height          =   195
      Left            =   1470
      TabIndex        =   5
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "DATE :"
      Height          =   195
      Left            =   1500
      TabIndex        =   2
      Top             =   930
      Width           =   1035
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_IncomeLain.frx":46B3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15210
   End
End
Attribute VB_Name = "frm_trans_IncomeLain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs2 As New ADODB.Recordset
Dim nomer As Integer
Public editTrans As Boolean
Public v_value As Double
Public v_nomer As Integer

Private Sub createKar()
   With LynxGrid2
      .AddColumn "EMP. CODE", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "EMP. NAME", 2000, , , , , , , , , True
      .AddColumn "DEPT. CODE", , , , , , , , , False
      .AddColumn "DEPARTMENT", 2000, , , , , , , , , True
      .AddColumn "DIV. CODE", , , , , , , , , False
      .AddColumn "SECTION", 2000, , , , , , , , , True
      .AddColumn "TITLE CODE", , , , , , , , , False
      .AddColumn "JOB TITLE", 2000, , , , , , , , , True
      .AddColumn "EMPLOYEE CODE", 2000, , , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub BuatKode()
Dim bulan As String

    strsql = "Select fn_buatkode(max(nomer)) nomer,year(curdate()) tahun,month(curdate()) bulan " _
        & "from t_income_expense " _
        & "WHERE userinput = '" & LOGIN_CODE & "' AND month(tglinput) = month(curdate()) " _
        & "AND year(tglinput) = year(curdate())"
    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs2!nomer) = False Then
        nomer = rs2!nomer
'        bulan = IIf(Len(rs2!bulan) = 1, "0" & rs2!bulan, rs2!bulan)
'        txtnotrans.Text = LOGIN_CODE & "/OEE/" & bulan & "/" & Right(rs2!tahun, 2) & "/" & rs2!nomer
    Else
        nomer = "1"
'        bulan = IIf(Len(month(Date)) = 1, "0" & month(Date), month(Date))
'        txtnotrans.Text = LOGIN_CODE & "/OEE/" & bulan & "/" & Right(year(Date), 2) & "/00001"
    End If
    rs2.Close
End Sub


Private Sub isiGridKar(pilihan As Integer)
Dim access As String

    If pilihan = 1 Then
        LynxGrid2.Clear
        If LOGIN_LEVEL = 100 Then
            SQL = "select a.nik,a.employee_name,a.department_code,d.department_name,a.division_code, " & _
                    "b.division_name,a.title_code,c.title_name,a.employee_code " & _
                 "from m_employee a join m_department d on a.department_code = d.department_code " & _
                    "join m_division b on a.division_code = b.division_code and a.department_code = b.department_code and a.company_code = b.company_code " & _
                    "join m_title c on a.title_code = c.title_code " & _
                  "WHERE flag_active <> 0 AND a.company_code = '" & TDBCombo_company.Text & "' " & _
                    "AND (nik LIKE '%" & txtnik.Text & "%' " & _
                        "OR employee_name LIKE '%" & txtnik.Text & "%')"
        Else
            SQL = "select a.nik,a.employee_name,a.department_code,d.department_name,a.division_code, " & _
                    "b.division_name,a.title_code,c.title_name,a.employee_code " & _
                  "from m_employee a join m_department d on a.department_code = d.department_code " & _
                    "join m_division b on a.division_code = b.division_code and a.department_code = b.department_code and a.company_code = b.company_code " & _
                    "join m_title c on a.title_code = c.title_code " & _
                  "WHERE flag_active <> 0 AND a.company_code = '" & TDBCombo_company.Text & "' " & _
                    "AND (nik LIKE '%" & txtnik.Text & "%' " & _
                        "OR employee_name LIKE '%" & txtnik.Text & "%') " & _
                    "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
        End If
                
        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid2.AddItem rs2!nik & vbTab & rs2!EMPLOYEE_NAME & vbTab & _
                rs2!DEPARTMENT_CODE & vbTab & rs2!department_name & vbTab & _
                rs2!DIVISION_CODE & vbTab & rs2!division_name & vbTab & rs2!title_code & vbTab & _
                rs2!title_name & vbTab & rs2!employee_code
                rs2.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs2.RecordCount = 1 Then
                rs2.MoveFirst
                txtkdkar.Text = rs2!employee_code
                txtnik.Text = rs2!nik
                txtnmkar.Text = rs2!EMPLOYEE_NAME
                txt_department_code.Text = rs2!DEPARTMENT_CODE
                txt_dep_name.Text = rs2!department_name
                txtkddiv.Text = rs2!DIVISION_CODE
                txtdivision.Text = rs2!division_name
                txtkdtitle.Text = rs2!title_code
                txttitle.Text = rs2!title_name
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
            txtkdkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 8)
            txtnmkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txt_department_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
            txt_dep_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 3)
            txtkddiv.Text = LynxGrid2.CellText(LynxGrid2.Row, 4)
            txtdivision.Text = LynxGrid2.CellText(LynxGrid2.Row, 5)
            txtkdtitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
            txttitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 7)
            txtjumlah.SetFocus
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    createKar
    
    If editTrans = False Then
        BuatKode
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_trans_IncomeLain = Nothing
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

Private Sub txtnik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err

    '++++++++++++++++++++++ Cek Input Expense Value ++++++++++++
        If Not IsNumeric(txtjumlah.Text) Then
            MsgBox "Value Must Be Numeric", vbExclamation, headerMSG
            txtjumlah.Text = ""
            txtjumlah.SetFocus
            Exit Sub
        End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++APAKAH KODE KARYAWAN SUDAH BENAR++++++++++++++++++++
    SQL = "SELECT 1 FROM m_employee WHERE employee_code = '" & txtkdkar.Text & "'"
    rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rs2.RecordCount = 0 Then
        MsgBox "Invalid Employee Code...!!", vbCritical
        Exit Sub
    End If
    rs2.Close
    
    CnG.BeginTrans
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If editTrans Then
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        If v_value <> txtjumlah.Text Then
            SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
            CnG.Execute SQL
        End If
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        SQL = "UPDATE t_income_expense set tgltrans = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'," _
                & "employee_code = '" & txtkdkar.Text & "',jmlpotong = " & Val(DropAllComma(txtjumlah)) & "," _
                & "useredit ='" & LOGIN_CODE & "',tgledit = now(),remark = '" & txtket.Text & "'," _
                & "flag_income_expense = " & Combo1.ListIndex & ",flag_type = '" & cboType.ListIndex & "',nm_type = '" & cboType.Text & "' " _
            & "WHERE employee_code = '" & txtkdkar.Text & "' and nomer = '" & v_nomer & "' " _
                & "AND flag_income_expense = '" & Combo1.ListIndex & "' AND flag_type = '" & cboType.ListIndex & "'"
    Else
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute SQL
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        SQL = "INSERT INTO t_income_expense (nomer,tgltrans,employee_code,jmlpotong,userinput,tglinput,remark,flag_income_expense,flag_type,nm_type) VALUES " _
                & "('" & nomer & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'," _
                & "'" & txtkdkar.Text & "'," & Val(DropAllComma(txtjumlah)) & ",'" & LOGIN_CODE & "'," _
                & "now(),'" & txtket.Text & "','" & Combo1.ListIndex & "','" & cboType.ListIndex & "','" & cboType.Text & "')"
    End If
    CnG.Execute SQL
    
    CnG.CommitTrans
    
    MsgBox "Save Successfully!", vbInformation, headerMSG
    
    If editTrans = True Then
        frm_List_IncomeLain.isiGridAbsen
        Unload Me
        Exit Sub
    Else
        BuatKode
        
        txtkdkar.Text = ""
        txtnik.Text = ""
        txtnmkar.Text = ""
        txtkddiv.Text = ""
        txtdivision.Text = ""
        txtkdtitle.Text = ""
        txttitle.Text = ""
        txtjumlah.Text = ""
        txtket.Text = ""
    End If
    
    DTPicker1.SetFocus
    
    frm_List_IncomeLain.isiGridAbsen
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub txtjumlah_Validate(Cancel As Boolean)
If Not Trim(txtjumlah) = "" Then
    txtjumlah = FormatNumber(DropAllComma(txtjumlah))
End If
End Sub

Private Sub txtjumlah_GotFocus()
    txtjumlah.SetFocus
    txtjumlah.SelStart = 0
    txtjumlah.SelLength = Len(txtjumlah.Text)
End Sub

Private Sub vbbutton2_Click()
    isiGridKar (1)
End Sub
