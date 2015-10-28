VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_UangMakan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PT KTC"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8010
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2850
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   15
      Top             =   1290
      Width           =   3855
   End
   Begin prj_fej_jkt.EnterNum EnterNum1 
      Height          =   285
      Left            =   1860
      TabIndex        =   13
      Top             =   1950
      Width           =   2355
      _ExtentX        =   4154
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
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   150
      TabIndex        =   12
      Top             =   3390
      Width           =   7725
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   0
      Top             =   2280
      Width           =   6165
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   5580
      TabIndex        =   7
      Top             =   540
      Width           =   2025
   End
   Begin VB.TextBox txtentry 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   5580
      TabIndex        =   5
      Top             =   210
      Width           =   2025
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   990
      TabIndex        =   3
      Top             =   270
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      Format          =   166526977
      CurrentDate     =   40794
   End
   Begin prj_fej_jkt.vbButton vbButton5 
      Height          =   585
      Left            =   2610
      TabIndex        =   2
      Top             =   3750
      Width           =   1575
      _ExtentX        =   2778
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
      MICON           =   "frm_trans_UangMakan.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_fej_jkt.vbButton vbButton1 
      Height          =   585
      Left            =   990
      TabIndex        =   1
      Top             =   3750
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1032
      BTYPE           =   14
      TX              =   "SAVE (F5)"
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
      MICON           =   "frm_trans_UangMakan.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1020
      OleObjectBlob   =   "frm_trans_UangMakan.frx":0038
      TabIndex        =   14
      Top             =   1290
      Width           =   1785
   End
   Begin VB.Label Label5 
      Caption         =   "Title :"
      Height          =   210
      Left            =   540
      TabIndex        =   11
      Top             =   1350
      Width           =   465
   End
   Begin VB.Label Label8 
      Caption         =   "Meals Allowance :"
      Height          =   195
      Left            =   210
      TabIndex        =   10
      Top             =   1980
      Width           =   1605
   End
   Begin VB.Label Label9 
      Caption         =   "Remark :"
      Height          =   210
      Left            =   360
      TabIndex        =   9
      Top             =   2340
      Width           =   645
   End
   Begin VB.Label Label11 
      Caption         =   "User Name"
      Height          =   210
      Left            =   4650
      TabIndex        =   8
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label10 
      Caption         =   "Entry Date"
      Height          =   210
      Left            =   4650
      TabIndex        =   6
      Top             =   270
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Date :"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   330
      Width           =   495
   End
End
Attribute VB_Name = "frm_trans_UangMakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim strSQL As String

Private Sub createKar()
   With LynxGrid2
      .AddColumn "NIK", 700, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "Div. code", , , , , , , , , False
      .AddColumn "Division", 2000, , , , , , , , , True
      .AddColumn "title code", , , , , , , , , False
      .AddColumn "Position", 2000, , , , , , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear
        strSQL = "select employee_code,employee_name,a.division_code,b.division_name," _
                    & "a.title_code,c.title_name " _
                & "from m_employee a join m_division b on a.division_code = b.division_code " _
                & "join m_title c on a.title_code = c.title_code " _
                & "WHERE a.company_code = '" & TDBCombo_company.Text & "' AND (a.employee_code LIKE '%" & txtkdkar.Text & "%' " _
                & "OR a.employee_name LIKE '%" & txtkdkar.Text & "%')"
                
        rs2.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid2.AddItem rs2!employee_code & vbTab & rs2!employee_name & vbTab & _
                rs2!division_code & vbTab & rs2!division_name & vbTab & rs2!title_code & vbTab & rs2!title_name
                rs2.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs2.RecordCount = 1 Then
                rs2.MoveFirst
                txtkdkar.Text = rs2!employee_code
                txtnmkar.Text = rs2!employee_name
                txtkddiv.Text = rs2!division_code
                txtdivision.Text = rs2!division_name
                txtkdtitle.Text = rs2!title_code
                txttitle.Text = rs2!title_name
                ttin.SetFocus
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else
            
        End If
        rs2.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txtkdkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txtnmkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txtkddiv.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
            txtdivision.Text = LynxGrid2.CellText(LynxGrid2.Row, 3)
            txtkdtitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 4)
            txttitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 5)
            ttin.SetFocus
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    isiTitle
    DTPicker1.Value = Date
End Sub

Private Sub isiTitle()
Dim rsTitle As New ADODB.Recordset
    strSQL = "select title_code,title_name " _
            & "from m_title order by 1"
    rsTitle.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo1.RowSource = rsTitle

End Sub

Private Sub vbButton1_Click()
    Me.MousePointer = vbHourglass
    
    '+++++++++++++++++++APAKAH KODE title SUDAH BENAR++++++++++++++++++++
    strSQL = "SELECT 1 FROM m_title WHERE title_code = '" & TDBCombo_company.Text & "'"
    rs2.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rs2.RecordCount = 0 Then
        MsgBox "Invalid Title Code...!!", vbCritical
        Exit Sub
    End If
    rs2.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++MENAMPILKAN DATA SALARY UNTUK DI UPDATE UANG MAKAN NYA +++++++++++++++++++
    strSQL = "SELECT a.employee_code," _
                & "(SELECT salary_date FROM m_salary WHERE employee_code = a.employee_code ORDER BY salary_date DESC LIMIT 1) salary_date " _
        & "FROM m_employee a "
    rs2.Open strSQL, CnG, adOpenDynamic, adLockReadOnly
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '======================UPDATE DATABASE====================================
    CnG.BeginTrans
    
    '++++++++++++++++++++++++++++++++++++
    strSQL = "UPDATE m_title SET uangmakan = '" & EnterNum1.Value & "',ketuangmakan = '" & txtket.Text & "' " _
        & "WHERE title_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute strSQL
    '++++++++++++++++++++++++++++++++++++
    
    rs2.MoveFirst
    While Not rs2.EOF
        If rs2!salary_date Is Null Then
            rs2.MoveNext
        Else
            strSQL = "UPDATE m_salary set uangmakan = '" & EnterNum1.Value & "' " _
                & "WHERE employee_code = '" & rs2!employee_code & "' AND salary_date = '" & rs2!salary_date & "'"
            CnG.Execute strSQL
        End If
    rs2.MoveNext
    Wend
    
    CnG.CommitTrans
    '==========================================================================
    Me.MousePointer = vbNormal
    EnterNum1.Value = 0
    txtket.Text = ""
End Sub

Private Sub vbButton2_Click()
    isiGridKar (1)
End Sub

Private Sub vbButton5_Click()
    Unload Me
End Sub
