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
   Begin prj_fej_jkt.LynxGrid LynxGrid2 
      Height          =   3735
      Left            =   1320
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   6588
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
      ColumnSort      =   -1  'True
   End
   Begin prj_fej_jkt.EnterNum txt_kali 
      Height          =   345
      Left            =   3450
      TabIndex        =   2
      Top             =   2220
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   609
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
   Begin prj_fej_jkt.EnterNum txtgaji_bersih 
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   2190
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
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
   Begin VB.TextBox txtstart_working 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1320
      TabIndex        =   34
      Top             =   1800
      Width           =   2205
   End
   Begin VB.TextBox txtdept_code 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   7020
      TabIndex        =   32
      Top             =   1110
      Width           =   885
   End
   Begin VB.TextBox txtsite_code 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   1320
      TabIndex        =   31
      Top             =   1110
      Width           =   885
   End
   Begin prj_fej_jkt.vbButton vbButton2 
      Height          =   285
      Left            =   2850
      TabIndex        =   29
      Top             =   780
      Width           =   345
      _ExtentX        =   609
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1320
      TabIndex        =   28
      Top             =   180
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   102432771
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
      TabIndex        =   27
      Top             =   1110
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   150
      TabIndex        =   26
      Top             =   4530
      Width           =   10695
   End
   Begin VB.TextBox txtkdtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   6540
      TabIndex        =   25
      Top             =   1470
      Width           =   885
   End
   Begin VB.TextBox txtkddiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1320
      TabIndex        =   24
      Top             =   1470
      Width           =   735
   End
   Begin VB.TextBox txtkdkar 
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
      TabIndex        =   19
      Top             =   780
      Width           =   4005
   End
   Begin VB.TextBox txtdivision 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2070
      TabIndex        =   18
      Top             =   1470
      Width           =   2535
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   7440
      TabIndex        =   17
      Top             =   1470
      Width           =   3195
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   6
      Top             =   3780
      Width           =   6075
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   15
      Top             =   480
      Width           =   2025
   End
   Begin VB.TextBox txtentry 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   13
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
      TabIndex        =   10
      Top             =   1110
      Width           =   3525
   End
   Begin prj_fej_jkt.vbButton vbButton5 
      Height          =   585
      Left            =   2790
      TabIndex        =   9
      Top             =   5160
      Width           =   1785
      _ExtentX        =   3149
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
   Begin prj_fej_jkt.EnterNum txt_jml_pesangon 
      Height          =   345
      Left            =   4470
      TabIndex        =   38
      Top             =   2220
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   609
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin prj_fej_jkt.EnterNum txt_other2 
      Height          =   345
      Left            =   4470
      TabIndex        =   5
      Top             =   3390
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   609
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
   Begin prj_fej_jkt.EnterNum txt_15_persen 
      Height          =   345
      Left            =   4470
      TabIndex        =   3
      Top             =   2610
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   609
      Value           =   15
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
   Begin prj_fej_jkt.vbButton vbButton1 
      Height          =   585
      Left            =   720
      TabIndex        =   7
      Top             =   5160
      Width           =   1785
      _ExtentX        =   3149
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
   Begin prj_fej_jkt.EnterNum txt_other1 
      Height          =   345
      Left            =   4470
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   609
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
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "%"
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
      Left            =   5220
      TabIndex        =   45
      Top             =   2670
      Width           =   225
   End
   Begin VB.Label Label19 
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
      Left            =   4200
      TabIndex        =   44
      Top             =   3450
      Width           =   225
   End
   Begin VB.Label Label18 
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
      Left            =   4200
      TabIndex        =   43
      Top             =   3060
      Width           =   225
   End
   Begin VB.Label Label17 
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
      Left            =   4200
      TabIndex        =   42
      Top             =   2670
      Width           =   225
   End
   Begin VB.Label Label16 
      Caption         =   "Other 2"
      Height          =   255
      Left            =   150
      TabIndex        =   41
      Top             =   3420
      Width           =   1815
   End
   Begin VB.Label Label15 
      Caption         =   "Other 1"
      Height          =   255
      Left            =   150
      TabIndex        =   40
      Top             =   3030
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "Med + Housing :"
      Height          =   255
      Left            =   150
      TabIndex        =   39
      Top             =   2640
      Width           =   1155
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
      Left            =   4200
      TabIndex        =   37
      Top             =   2280
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
      Left            =   3180
      TabIndex        =   36
      Top             =   2280
      Width           =   225
   End
   Begin VB.Label Label8 
      Caption         =   "Pension Value :"
      Height          =   255
      Left            =   150
      TabIndex        =   35
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "Start Working :"
      Height          =   210
      Left            =   210
      TabIndex        =   33
      Top             =   1830
      Width           =   1065
   End
   Begin VB.Label Label5 
      Caption         =   "Employee :"
      Height          =   210
      Left            =   480
      TabIndex        =   23
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label6 
      Caption         =   "Division :"
      Height          =   210
      Left            =   630
      TabIndex        =   22
      Top             =   1500
      Width           =   675
   End
   Begin VB.Label Label7 
      Caption         =   "Title :"
      Height          =   210
      Left            =   6090
      TabIndex        =   21
      Top             =   1530
      Width           =   405
   End
   Begin VB.Label Label9 
      Caption         =   "Remark :"
      Height          =   210
      Left            =   150
      TabIndex        =   20
      Top             =   3840
      Width           =   645
   End
   Begin VB.Label Label11 
      Caption         =   "User Name"
      Height          =   210
      Left            =   7830
      TabIndex        =   16
      Top             =   540
      Width           =   885
   End
   Begin VB.Label Label10 
      Caption         =   "Entry Date"
      Height          =   210
      Left            =   7830
      TabIndex        =   14
      Top             =   210
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Department :"
      Height          =   195
      Left            =   6060
      TabIndex        =   12
      Top             =   1170
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Branch Office :"
      Height          =   195
      Left            =   210
      TabIndex        =   11
      Top             =   1170
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Date :"
      Height          =   195
      Left            =   510
      TabIndex        =   8
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
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
Dim access As String

access = IIf(LOGIN_LEVEL = 100, "", "AND (a.managerial_access = 0 OR a.managerial_access IS NULL)")
    
    If pilihan = 1 Then
        LynxGrid2.Clear
        strsql = "SELECT a.employee_code,a.employee_name," _
                    & "a.company_code,e.company_name," _
                    & "a.department_code,d.department_name," _
                    & "a.division_code,b.division_name," _
                    & "a.title_code,c.title_name,a.start_working," _
                    & "(select salary from m_salary where employee_code = a.employee_code order by salary_date desc limit 1) salary " _
                & "FROM m_employee a LEFT JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                & "LEFT JOIN m_title c ON a.title_code = c.title_code " _
                & "LEFT JOIN m_department d ON a.department_code = d.department_code and d.company_code = a.company_code " _
                & "JOIN m_company e ON a.company_code = e.company_code " _
                & "WHERE (a.employee_code LIKE '%" & txtkdkar.Text & "%' " _
                & "OR a.employee_name LIKE '%" & txtkdkar.Text & "%') " _
                & "AND a.flag_active <> 0 " & access & " " _
                & "ORDER BY a.company_code, a.department_code, a.division_code, a.employee_name ASC"
                
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid2.AddItem rs2!employee_code & vbTab & rs2!employee_name & vbTab & _
                    rs2!COMPANY_CODE & vbTab & rs2!company_name & vbTab & rs2!department_code & vbTab & rs2!department_name & vbTab & _
                    rs2!division_code & vbTab & rs2!division_name & vbTab & rs2!title_code & vbTab & _
                    rs2!title_name & vbTab & rs2!start_working & vbTab & rs2!salary
                rs2.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs2.RecordCount = 1 Then
                rs2.MoveFirst
                txtkdkar.Text = rs2!employee_code
                txtnmkar.Text = rs2!employee_name
                txtsite_code.Text = rs2!COMPANY_CODE
                txt_company_name.Text = rs2!company_name
                txtdept_code.Text = rs2!department_code
                txtnmDept.Text = rs2!department_name
                txtkddiv.Text = rs2!division_code
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
            txtkdkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
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
        
        Dim rsincome As New ADODB.Recordset
        Dim v_masa_kerja As Double
        Dim v_wisdom As Double
        Dim v_appreciation As Double
        Dim v_housing As Double
        Dim v_leave As Double
        Dim v_leave_value As Double
        
        strsql = "SELECT ((YEAR(Now()) - YEAR(a.start_working))) + (RIGHT(Now(),5) >= RIGHT(DATE(a.start_working),5)) AS masa_kerja," _
            & "(SELECT over_leave FROM t_leave_periode WHERE employee_code = a.employee_code ORDER BY start_periode, end_periode DESC LIMIT 1) * -1 AS sisa_cuti " _
            & "FROM m_employee a WHERE a.employee_code = '" & txtkdkar.Text & "'"
        rsincome.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsincome.RecordCount > 0 Then
            v_masa_kerja = rsincome!masa_kerja
            v_wisdom = 2 * txtgaji_bersih.Value * v_masa_kerja
            v_appreciation = txt_kali.Value * txtgaji_bersih.Value
            v_housing = (txt_15_persen.Value / 100) * (v_wisdom + v_appreciation)
            v_leave = rsincome!sisa_cuti
            v_leave_value = (txtgaji_bersih.Value / 25) * v_leave
        End If
        rsincome.Close
        
        v_loan = IIf(IsNull(rsproses!loan), 0, rsproses!loan)
        v_jamsostek = txtgaji_bersih.Value * 0.02
        v_others = IIf(IsNull(rsproses!potongan_lain), 0, rsproses!potongan_lain)
        v_absen = (rsproses!jmlHariAbsent / 25) * txtgaji_bersih.Value
        v_absen_others = (v_others + v_absen)
        
        v_tot_income = (v_wisdom + v_appreciation + v_housing + v_leave_value + txt_other1.Value + txt_other2.Value)
        
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_comp WHERE pph21_number = 1"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If v_tot_income > rspph!pph21_upper Then '50000000
                v_pph_0 = (rspph!pph21_percentage / 100) * rspph!pph21_upper '50000000
            Else
                v_pph_0 = (rspph!pph21_percentage / 100) * v_tot_income
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_comp WHERE pph21_number = 2"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If v_tot_income <= rspph!pph21_under Then '50000000
                v_pph_5 = 0
            ElseIf v_tot_income < rspph!pph21_upper Then '100000000
                v_pph_5 = (rspph!pph21_percentage / 100) * (v_tot_income - rspph!pph21_under) '50000000
            Else
'                v_pph_5 = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '50000000
                v_pph_5 = (rspph!pph21_percentage / 100) * (v_tot_income - rspph!pph21_under) '50000000
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_comp WHERE pph21_number = 3"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If v_tot_income <= rspph!pph21_under Then '500000000
                v_pph_15 = 0
            ElseIf v_tot_income < rspph!pph21_upper Then  '500000000
                v_pph_15 = (rspph!pph21_percentage / 100) * (v_tot_income - rspph!pph21_under) '100000000
            Else
                v_pph_15 = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '100000000
            End If
        End If
        rspph.Close
            
        strsql = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_comp WHERE pph21_number = 4"
        rspph.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rspph.RecordCount > 0 Then
            If v_tot_income <= rspph!pph21_under Then '500000000
                v_pph_25 = 0
            Else
                If v_tot_income > rspph!pph21_under Then '500000000
                    v_pph_25 = (rspph!pph21_percentage / 100) * (v_tot_income - rspph!pph21_under) '500000000
                Else
                    v_pph_25 = 0
                End If
            End If
        End If
        rspph.Close
        
        v_pph21_pension = (v_pph_0 + v_pph_5 + v_pph_15 + v_pph_25)
        
        v_tot_deduction = (v_pph21_pension + v_loan + v_jamsostek + v_absen_others)
        v_total = (v_tot_income - v_tot_deduction)
    End If
    rsproses.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    If editTrans = False Then
        strsql = "INSERT INTO t_pensiun " & _
                "(tgltrans,employee_code,basic_salary,kali_gaji,uang_pensiun,persen_transport, " & _
                "lain1,lain2,ket,userinput,tglinput, " & _
                "pph21,loan,jamsostek,absen_others,total_income,total_deduction,total) " & _
                "VALUES " & _
                "('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & txtkdkar.Text & "'," & _
                "'" & txtgaji_bersih.Value & "','" & txt_kali.Value & "','" & txt_jml_pesangon.Value & "'," & _
                "'" & txt_15_persen.Value & "','" & txt_other1.Value & "','" & txt_other2.Value & "'," & _
                "'" & txtket.Text & "','" & LOGIN_CODE & "',now(), " & _
                "'" & v_pph21_pension & "','" & v_loan & "','" & v_jamsostek & "','" & v_absen_others & "'," & _
                "'" & v_tot_income & "','" & v_tot_deduction & "','" & v_total & "')"
    Else
        strsql = "UPDATE t_pensiun SET " & _
                "tgltrans = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',employee_code = '" & txtkdkar.Text & "'," & _
                "basic_salary = '" & txtgaji_bersih.Value & "',kali_gaji = '" & txt_kali.Value & "'," & _
                "uang_pensiun = '" & txt_jml_pesangon.Value & "',persen_transport = '" & txt_15_persen.Value & "'," & _
                "lain1 = '" & txt_other1.Value & "',lain2 = '" & txt_other2.Value & "'," & _
                "ket = '" & txtket.Text & "',useredit = '" & LOGIN_CODE & "',tgledit = now(), " & _
                "pph21 = '" & v_pph21_pension & "',loan = '" & v_loan & "',jamsostek = '" & v_jamsostek & "'," & _
                "absen_others = '" & v_absen_others & "',total_income = '" & v_tot_income & "'," & _
                "total_deduction = '" & v_tot_deduction & "',total = '" & v_total & "' " & _
                "WHERE employee_code = '" & txtkdkar.Text & "'"
    End If
    
    CnG.Execute strsql
    
    Dim clsinsert As New clsInsert_h_salary
    Dim bulan As String

    strsql = "DELETE FROM h_salary_new WHERE employee_code = '" & txtkdkar.Text & "' " _
            & "AND month = '" & Mid(Format(DTPicker1.Value, "yyyy-MM-dd"), 1, 7) & "'"
    CnG.Execute (strsql)

    bulan = Mid(Format(DTPicker1.Value, "yyyy-MM-dd"), 1, 7)

    strsql = "Select employee_code, employee_name," & _
            "marital_status, number_of_children, sex," & _
            "end_working, start_mc, CAST(IFNULL(end_mc,LAST_DAY('" & tgl & "')) as DATE) end_mc, flag_active, " & _
            "CONCAT(DATE_FORMAT(LAST_DAY(NOW()),'%Y-%m-'),'01') tgl_awal " & _
        "FROM m_employee " & _
        "WHERE employee_code = '" & txtkdkar.Text & "'"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

    Call clsinsert.Insert_h_salary(rs!employee_code, rs!sex, bulan, Format(rs!tgl_awal, "yyyy-MM-dd"), _
                Format(DTPicker1.Value, "yyyy-MM-dd"), rs!marital_status, IIf(IsNull(rs!number_of_children), 0, rs!number_of_children), _
                IIf(IsNull(Format(rs!start_mc, "yyyyMM")), "0", Format(rs!start_mc, "yyyyMM")), _
                rs!flag_active, txtsite_code.Text)

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
    txtkdkar.SetFocus
    
    MsgBox "Save Succesfully!", vbInformation, "Success!"
    
    Unload Me
    frm_List_pensiun.isiGridAbsen
End Sub

Private Sub vbButton2_Click()
    isiGridKar (1)
End Sub

Private Sub vbButton5_Click()
    Unload Me
End Sub
