VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_PotongLain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OTHER EMPLOYEE EXPENSE"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtkdkar 
      Height          =   345
      Left            =   4590
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   405
   End
   Begin prj_panji.vbButton vbButton2 
      Height          =   285
      Left            =   2550
      TabIndex        =   26
      Top             =   1350
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
      MICON           =   "frm_trans_PotongLain.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.LynxGrid LynxGrid2 
      Height          =   4545
      Left            =   1410
      TabIndex        =   25
      Top             =   1650
      Visible         =   0   'False
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   8017
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
   Begin VB.TextBox txtnotrans 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   8640
      TabIndex        =   24
      Top             =   180
      Width           =   2295
   End
   Begin VB.TextBox txt_dep_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2460
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   22
      Top             =   930
      Width           =   3855
   End
   Begin VB.PictureBox txtjumlah1 
      Height          =   345
      Left            =   5550
      ScaleHeight     =   285
      ScaleWidth      =   2655
      TabIndex        =   21
      Top             =   1980
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   330
      TabIndex        =   19
      Top             =   4530
      Width           =   10695
   End
   Begin VB.TextBox txtkdtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1410
      TabIndex        =   18
      Top             =   1650
      Width           =   795
   End
   Begin VB.TextBox txtkddiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   7800
      TabIndex        =   17
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtnik 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1410
      TabIndex        =   0
      Top             =   1350
      Width           =   1125
   End
   Begin VB.TextBox txtnmkar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2910
      TabIndex        =   12
      Top             =   1350
      Width           =   4005
   End
   Begin VB.TextBox txtdivision 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8550
      TabIndex        =   11
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2220
      TabIndex        =   10
      Top             =   1650
      Width           =   3195
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2730
      Width           =   9015
   End
   Begin VB.ComboBox cmbdep 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1410
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   930
      Width           =   1005
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   570
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1410
      TabIndex        =   2
      Top             =   180
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   153157635
      CurrentDate     =   40794
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1410
      OleObjectBlob   =   "frm_trans_PotongLain.frx":001C
      TabIndex        =   5
      Top             =   570
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
      Height          =   345
      Left            =   1500
      TabIndex        =   23
      Top             =   2340
      Width           =   2655
   End
   Begin prj_panji.vbButton vbButton1 
      Height          =   585
      Left            =   1140
      TabIndex        =   27
      Top             =   5160
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1032
      BTYPE           =   2
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
      MICON           =   "frm_trans_PotongLain.frx":1FDA
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
      Height          =   585
      Left            =   2730
      TabIndex        =   28
      Top             =   5160
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1032
      BTYPE           =   2
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
      MICON           =   "frm_trans_PotongLain.frx":1FF6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frm_trans_PotongLain.frx":2012
      Left            =   1620
      List            =   "frm_trans_PotongLain.frx":201F
      TabIndex        =   30
      Top             =   1980
      Width           =   2925
   End
   Begin VB.Label Label11 
      Caption         =   "* yyyy-MM-dd"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3240
      TabIndex        =   32
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Label4 
      Caption         =   "Tipe Pengeluaran:"
      Height          =   210
      Left            =   240
      TabIndex        =   31
      Top             =   2070
      Width           =   1785
   End
   Begin VB.Label Label8 
      Caption         =   "Jumlah :"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   2430
      Width           =   1185
   End
   Begin VB.Label Label5 
      Caption         =   "Karyawan :"
      Height          =   210
      Left            =   240
      TabIndex        =   16
      Top             =   1380
      Width           =   795
   End
   Begin VB.Label Label6 
      Caption         =   "Divisi :"
      Height          =   210
      Left            =   7110
      TabIndex        =   15
      Top             =   1380
      Width           =   675
   End
   Begin VB.Label Label7 
      Caption         =   "Jabatan :"
      Height          =   210
      Left            =   240
      TabIndex        =   14
      Top             =   1710
      Width           =   675
   End
   Begin VB.Label Label9 
      Caption         =   "Keterangan :"
      Height          =   210
      Left            =   240
      TabIndex        =   13
      Top             =   2790
      Width           =   1005
   End
   Begin VB.Label Label10 
      Caption         =   "No. Transaksi :"
      Height          =   210
      Left            =   7440
      TabIndex        =   9
      Top             =   210
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Dept. / Area :"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Perusahaan :"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   630
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal :"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frm_trans_PotongLain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs2 As New ADODB.Recordset
Dim strsql As String
Dim nomer As Integer
Public editTrans As Boolean
Public v_value As Double

Private Sub createKar()
   With LynxGrid2
      .AddColumn "NIK", 700, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "Div. code", , , , , , , , , False
      .AddColumn "Division", 2000, , , , , , , , , True
      .AddColumn "title code", , , , , , , , , False
      .AddColumn "Position", 2000, , , , , , , , , True
      .AddColumn "Emp Code", 2000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub BuatKode()
Dim bulan As String

    strsql = "Select fn_buatkode(max(nomer)) nomer,year(curdate()) tahun,month(curdate()) bulan " _
        & "from t_employee_expense " _
        & "WHERE userinput = '" & LOGIN_CODE & "' AND month(tglinput) = month(curdate()) " _
        & "AND year(tglinput) = year(curdate())"
    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs2!nomer) = False Then
        nomer = rs2!nomer
        bulan = IIf(Len(rs2!bulan) = 1, "0" & rs2!bulan, rs2!bulan)
        txtnotrans.Text = LOGIN_CODE & "/OEE/" & bulan & "/" & Right(rs2!tahun, 2) & "/" & rs2!nomer
    Else
        nomer = "1"
        bulan = IIf(Len(month(Date)) = 1, "0" & month(Date), month(Date))
        txtnotrans.Text = LOGIN_CODE & "/OEE/" & bulan & "/" & Right(year(Date), 2) & "/00001"
    End If
    rs2.Close
End Sub

Private Sub isiGridKar(pilihan As Integer)
Dim access As String

    If pilihan = 1 Then
        LynxGrid2.Clear
        strsql = "select nik,employee_name,a.division_code,b.division_name," _
                    & "a.title_code,c.title_name,employee_code " _
                & "from m_employee a join m_division b on a.division_code = b.division_code and a.company_code = b.company_code and a.department_code = b.department_code " _
                & "join m_title c on a.title_code = c.title_code and a.company_code = c.company_code " _
                & "WHERE a.company_code = '" & TDBCombo_company.Text & "' AND a.department_code = '" & cmbdep.Text & "' AND " _
                & "(a.nik LIKE '%" & txtnik.Text & "%' " _
                & "OR a.employee_name LIKE '%" & txtnik.Text & "%') AND " _
                & "(level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) ORDER BY employee_name"
        
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid2.AddItem rs2!nik & vbTab & rs2!EMPLOYEE_NAME & vbTab & _
                rs2!division_code & vbTab & rs2!division_name & vbTab & rs2!title_code & vbTab & rs2!title_name & vbTab & rs2!employee_code
                rs2.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs2.RecordCount = 1 Then
                rs2.MoveFirst
                txtkdkar.Text = rs2!employee_code
                txtnik.Text = rs2!nik
                txtnmkar.Text = rs2!EMPLOYEE_NAME
                txtkddiv.Text = rs2!division_code
                txtdivision.Text = rs2!division_name
                txtkdtitle.Text = rs2!title_code
                txttitle.Text = rs2!title_name
                txtjumlah.SetFocus
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else
            
        End If
        rs2.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txtkdkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
            txtnik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txtnmkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txtkddiv.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
            txtdivision.Text = LynxGrid2.CellText(LynxGrid2.Row, 3)
            txtkdtitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 4)
            txttitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 5)
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
    
'    txtjumlah.Text = Format(0, "#,##0.00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_trans_PotongLain = Nothing
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

Private Sub vbButton1_Click()
On Error GoTo Err

Dim start_time As String, end_time As String, max_break_out As String, min_break_in As String
Dim time_in As String, time_out_break As String, time_in_break As String, time_out As String


    '++++++++++++++++++++++ Cek Input Expense Value ++++++++++++
        If Not IsNumeric(txtjumlah.Text) Then
            MsgBox "Expense Value Must Be Numeric", vbExclamation, headerMSG
            txtjumlah.Text = ""
            txtjumlah.SetFocus
            Exit Sub
        End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    '+++++++++++++++++++APAKAH KODE KARYAWAN SUDAH BENAR++++++++++++++++++++
    strsql = "SELECT 1 FROM m_employee WHERE employee_code = '" & txtkdkar.Text & "'"
    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs2.RecordCount = 0 Then
        MsgBox "Kode Karyawan Tidak Valid...", vbCritical
        Exit Sub
    End If
    rs2.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If editTrans Then
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        If v_value <> txtjumlah.Text Then
            strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
            CnG.Execute strsql
        End If
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        strsql = "UPDATE t_employee_expense set tgltrans = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'," _
                & "employee_code = '" & txtkdkar.Text & "',jmlpotong = " & Val(DropAllComma(txtjumlah)) & "," _
                & "useredit ='" & LOGIN_CODE & "',tgledit = now(),remark = '" & txtket.Text & "',flag_other_expense = '" & Combo1.ListIndex & "' " _
            & "WHERE notrans = '" & txtnotrans.Text & "'"
    Else
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute strsql
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        strsql = "INSERT INTO t_employee_expense (nomer,notrans,tgltrans,employee_code,jmlpotong,userinput,tglinput,remark,flag_other_expense) VALUES " _
                & "('" & nomer & "','" & txtnotrans.Text & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'," _
                & "'" & txtkdkar.Text & "'," & Val(DropAllComma(txtjumlah)) & ",'" & LOGIN_CODE & "',now(),'" & txtket.Text & "','" & Combo1.ListIndex & "')"
    End If
    CnG.Execute strsql
    
    If editTrans Then
        frm_List_PotonganLain.isiGridAbsen
        Unload Me
        Exit Sub
    End If
    
    MsgBox "Save Succesfully!", vbInformation, headerMSG
    
    BuatKode
    txtkdkar.Text = ""
    txtnik.Text = ""
    txtnmkar.Text = ""
    txtjumlah.Text = ""
    txtket.Text = ""
    DTPicker1.SetFocus
    
    frm_List_PotonganLain.isiGridAbsen
    Exit Sub

Err:
'MsgBox Err.Number
If Err.Number = -2147217900 Then
    MsgBox "No. Trans Already Exist!" & Chr(13) & _
        "Please Change No. Trans..", vbExclamation, headerMSG
Else
    MsgBox Err.Description
End If
Exit Sub

End Sub

Private Sub vbbutton2_Click()
    isiGridKar (1)
End Sub

Private Sub vbButton5_Click()
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
