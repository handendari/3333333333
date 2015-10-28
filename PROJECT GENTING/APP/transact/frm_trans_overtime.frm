VERSION 5.00
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_overtime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manual Overtime"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtkdkar 
      Height          =   315
      Left            =   4500
      TabIndex        =   39
      Top             =   270
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtjmljam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   35.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   5220
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Hari Libur"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1290
      TabIndex        =   1
      Top             =   2190
      Width           =   1005
   End
   Begin VB.TextBox Txt4jam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   35.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   9750
      TabIndex        =   38
      Top             =   2130
      Width           =   1125
   End
   Begin VB.TextBox Txt3jam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   35.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8310
      TabIndex        =   36
      Top             =   2130
      Width           =   1095
   End
   Begin VB.TextBox txt2jam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   35.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6840
      TabIndex        =   34
      Top             =   2130
      Width           =   1095
   End
   Begin VB.TextBox txt15jam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   35.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5370
      TabIndex        =   32
      Top             =   2130
      Width           =   1095
   End
   Begin VB.CommandButton vbbutton5 
      Caption         =   "EXIT"
      Height          =   585
      Left            =   2790
      TabIndex        =   30
      Top             =   5160
      Width           =   1905
   End
   Begin VB.CommandButton vbbutton2 
      Caption         =   "..."
      Height          =   285
      Left            =   2430
      TabIndex        =   29
      Top             =   1440
      Width           =   315
   End
   Begin VB.CommandButton vbbutton1 
      Caption         =   "SAVE (F5)"
      Height          =   585
      Left            =   750
      TabIndex        =   4
      Top             =   5160
      Width           =   1905
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1320
      TabIndex        =   27
      Top             =   180
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   95158275
      CurrentDate     =   40823
   End
   Begin VB.TextBox txtnmDept 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   26
      Top             =   930
      Width           =   2805
   End
   Begin VB.TextBox txtkdshift 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6660
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   25
      Top             =   930
      Width           =   855
   End
   Begin VB.TextBox txtnmshift 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   7530
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   24
      Top             =   930
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   150
      TabIndex        =   22
      Top             =   4620
      Width           =   10695
   End
   Begin VB.TextBox txtkdtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1290
      TabIndex        =   21
      Top             =   1770
      Width           =   795
   End
   Begin VB.TextBox txtkddiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   7590
      TabIndex        =   20
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txt_nik 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1290
      TabIndex        =   0
      Top             =   1440
      Width           =   1125
   End
   Begin VB.TextBox txtnmkar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2790
      TabIndex        =   14
      Top             =   1440
      Width           =   4005
   End
   Begin VB.TextBox txtdivision 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8340
      TabIndex        =   13
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2100
      TabIndex        =   12
      Top             =   1770
      Width           =   3195
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   3
      Top             =   3180
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   10
      Top             =   480
      Width           =   2025
   End
   Begin VB.TextBox txtentry 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   8
      Top             =   150
      Width           =   2025
   End
   Begin VB.ComboBox cmbdep 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1290
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   930
      Width           =   1095
   End
   Begin TDBTime6Ctl.TDBTime ttin 
      Height          =   285
      Left            =   3630
      TabIndex        =   40
      Top             =   2160
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frm_trans_overtime.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frm_trans_overtime.frx":006C
      Spin            =   "frm_trans_overtime.frx":00BC
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "00:00"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   0
   End
   Begin TDBTime6Ctl.TDBTime ttout 
      Height          =   285
      Left            =   3630
      TabIndex        =   41
      Top             =   2490
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frm_trans_overtime.frx":00E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frm_trans_overtime.frx":0150
      Spin            =   "frm_trans_overtime.frx":01A0
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "00:00"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   0
   End
   Begin prj_genting.LynxGrid LynxGrid2 
      Height          =   4515
      Left            =   1290
      TabIndex        =   31
      Top             =   1740
      Visible         =   0   'False
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   7964
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
   Begin VB.Label Label16 
      Caption         =   "* yyyy-MM-dd"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2940
      TabIndex        =   43
      Top             =   210
      Width           =   1425
   End
   Begin VB.Label Label15 
      Caption         =   "Jam Lembur Out:"
      Height          =   195
      Left            =   2460
      TabIndex        =   42
      Top             =   2520
      Width           =   1125
   End
   Begin VB.Label Label14 
      Caption         =   "4 X"
      Height          =   210
      Left            =   9450
      TabIndex        =   37
      Top             =   2190
      Width           =   285
   End
   Begin VB.Label Label12 
      Caption         =   "3 X"
      Height          =   210
      Left            =   8010
      TabIndex        =   35
      Top             =   2190
      Width           =   285
   End
   Begin VB.Label Label3 
      Caption         =   "2 X"
      Height          =   210
      Left            =   6540
      TabIndex        =   33
      Top             =   2190
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "1.5 X"
      Height          =   210
      Left            =   4890
      TabIndex        =   28
      Top             =   2190
      Width           =   435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Shift :"
      Height          =   195
      Left            =   6210
      TabIndex        =   23
      Top             =   990
      Width           =   405
   End
   Begin VB.Label Label5 
      Caption         =   "Karyawan :"
      Height          =   210
      Left            =   450
      TabIndex        =   19
      Top             =   1500
      Width           =   795
   End
   Begin VB.Label Label6 
      Caption         =   "Divisi :"
      Height          =   210
      Left            =   7020
      TabIndex        =   18
      Top             =   1500
      Width           =   555
   End
   Begin VB.Label Label7 
      Caption         =   "Jabatan :"
      Height          =   210
      Left            =   600
      TabIndex        =   17
      Top             =   1830
      Width           =   675
   End
   Begin VB.Label Label8 
      Caption         =   "Jam Lembur In:"
      Height          =   195
      Left            =   2460
      TabIndex        =   16
      Top             =   2190
      Width           =   1125
   End
   Begin VB.Label Label9 
      Caption         =   "Remark :"
      Height          =   210
      Left            =   630
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label11 
      Caption         =   "Nama User"
      Height          =   210
      Left            =   7890
      TabIndex        =   11
      Top             =   540
      Width           =   825
   End
   Begin VB.Label Label10 
      Caption         =   "Tanggal Input"
      Height          =   210
      Left            =   7680
      TabIndex        =   9
      Top             =   210
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Dept / Area :"
      Height          =   195
      Left            =   330
      TabIndex        =   7
      Top             =   990
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal :"
      Height          =   195
      Left            =   780
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frm_trans_overtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs2 As New ADODB.Recordset
Dim strsql As String
Public editTrans As Boolean
Public v_ot_hour As Double
Dim v_tot_ot, v_15jam, v_2jam, v_3jam, v_4jam, v_tot_overtime As Double
Dim v_15, v_2, v_3, v_4 As Double
Dim v_holiday As Integer
Public v_tt_in As String
Public v_tt_out As String

Private Sub createKar()
   With LynxGrid2
      .AddColumn "NIK", 700, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "Div. code", , , , , , , , , False
      .AddColumn "Division", 2000, , , , , , , , , True
      .AddColumn "title code", , , , , , , , , False
      .AddColumn "Title", 2000, , , , , , , , , True
      .AddColumn "level", 2000, , , , , , , , False
      .AddColumn "Employee Code", 2000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear
        strsql = "select a.nik,a.employee_name,a.division_code,a.division_name," _
                    & "a.title_code,a.title_name,a.level_code,a.employee_code " _
                & "from m_employee a join m_salary_standard b on a.employee_code = b.employee_code " _
                & "WHERE flag_active <> 0 AND department_code = '" & cmbdep.Text & "' AND " _
                & "ifnull(b.flag_ot,0) = 1 AND " _
                & "(a.employee_code LIKE '%" & txtkdkar.Text & "%' " _
                & "OR a.employee_name LIKE '%" & txtkdkar.Text & "%') " _
                & "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
                
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid2.AddItem rs2!nik & vbTab & rs2!EMPLOYEE_NAME & vbTab & _
                rs2!division_code & vbTab & rs2!division_name & vbTab & rs2!title_code & vbTab & _
                rs2!title_name & vbTab & rs2!employee_code
                rs2.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs2.RecordCount = 1 Then
                rs2.MoveFirst
                txtkdkar.Text = rs2!employee_code
                txt_nik.Text = rs2!nik
                txtnmkar.Text = rs2!EMPLOYEE_NAME
                txtkddiv.Text = rs2!division_code
                txtdivision.Text = IIf(IsNull(rs2!division_name) = True, "", rs2!division_name)
                txtkdtitle.Text = rs2!title_code
                txttitle.Text = rs2!title_name
                Check1.SetFocus
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else
            
        End If
        rs2.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txt_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txtnmkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txtkddiv.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
            txtdivision.Text = LynxGrid2.CellText(LynxGrid2.Row, 3)
            txtkdtitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 4)
            txttitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 5)
            txtkdkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
            Check1.SetFocus
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub Combo1_Change()
    Call totalJam
End Sub

Private Sub Combo1_Click()
    Call totalJam
End Sub

'Private Sub totalJam()
'    Dim tglin As Date, tglout As Date
'    Dim jambreak As Integer
'
'    tglin = Date & " " & ttin.Value
'    If ttin.Value > ttout.Value Then
'        tglout = Date + 1 & " " & ttout.Value
'    Else
'        tglout = Date & " " & ttout.Value
'    End If
'    Select Case Combo1.ListIndex
'        Case 0
'            jambreak = 1
'        Case 1
'            jambreak = 2
'        Case 3
'            jambreak = 0
'    End Select
'    txtjmljam.Text = DateDiff("h", tglin, tglout) - jambreak
'End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        ttout.SetFocus
    End If
End Sub

Private Sub Combo3_Click()
    If Combo3.Text = "To" Then
        DTPicker3.Value = DTPicker2.Value
        DTPicker3.Visible = True
    Else
        DTPicker3.Visible = False
    End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then KeyAscii = 0
End Sub

Private Sub Check1_Click()
    txtjmljam_Change
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 116 Then
        vbButton1_Click
    End If
End Sub

Private Sub Form_Load()
Dim rsshift As New ADODB.Recordset

    createKar
    
    
'    If editTrans = False Then
'        isijamkerja
'    End If
        
End Sub

'Public Sub isijamkerja()
'    If WeekdayName(Weekday(frm_list_manual_att.DTPicker1.Value)) = "Friday" Then
'        Combo1.ListIndex = 1
'    End If
'
'End Sub
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

Private Sub TDBCombo1_ItemChange()
    If TDBCombo1.ApproxCount > 0 Then
        TDBCombo1.Text = TDBCombo1.Columns("absent_code").Value
        txtnmabsentstatus.Text = TDBCombo1.Columns("absent_name").Value
        
        If TDBCombo1.Text <> "0" Then
            DTPicker2.Visible = True
            Label16.Visible = True
            Combo3.Visible = True
            ttin.Enabled = False
            Combo1.Enabled = False
            ttout.Enabled = False
        Else
            Label16.Visible = False
            DTPicker2.Visible = False
            DTPicker3.Visible = False
            Combo3.Visible = False
            ttin.Enabled = True
            ttout.Enabled = True
            Combo1.Enabled = True
        End If
    End If
End Sub

Private Sub ttin_Change()
    Call totalJam
End Sub

Private Sub ttin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
End Sub

Private Sub ttout_Change()
    Call totalJam
End Sub

Private Sub ttout_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtket.SetFocus
    End If
End Sub

Private Sub txtjmljam_Change()
'    If Check1.Value = 1 Then
'        Select Case Val(txtjmljam.Text)
'            Case Is <= 7
'                txt15jam.Text = 0
'                txt2jam.Text = Val(txtjmljam.Text)
'                Txt3jam.Text = 0
'                Txt4jam.Text = 0
'            Case Is = 8
'                txt15jam.Text = 0
'                txt2jam.Text = 7
'                Txt3jam.Text = 1
'            Case Else
'                txt15jam.Text = 0
'                txt2jam.Text = 7
'                Txt3jam.Text = 1
'                Txt4jam.Text = Val(txtjmljam.Text) - Val(txt2jam.Text) - Val(Txt3jam.Text)
'        End Select
'    Else
'        Select Case Val(txtjmljam.Text)
'            Case Is = 1
'                txt15jam.Text = 1
'            Case Is >= 2
'                txt15jam.Text = 1
'                txt2jam.Text = Val(txtjmljam.Text) - Val(txt15jam.Text)
'                Txt3jam.Text = 0
'                Txt4jam.Text = 0
'            Case Else
'                txt15jam.Text = 0
'                txt2jam.Text = 0
'                Txt3jam.Text = 0
'                Txt4jam.Text = 0
'        End Select
'    End If
Call totalJam
End Sub

Private Sub txtkdkar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        vbButton1.SetFocus
    End If
End Sub

Private Sub vbButton1_Click()
'    If TDBCombo1.Text = "0" Then
'        If ttin.Value = ttout.Value Then
'            MsgBox "Invalid Check In Time...!!!", vbExclamation, "Error"
'            Exit Sub
'        End If
'    End If
'
    '+++++++++++++++++++APAKAH KODE KARYAWAN SUDAH BENAR++++++++++++++++++++
    strsql = "SELECT 1 FROM m_employee WHERE employee_code = '" & txtkdkar.Text & "'"
    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs2.RecordCount = 0 Then
        MsgBox "Invalid Employee Code...!!", vbCritical
        rs2.Close
        Exit Sub
    End If
    rs2.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++APAKAH KODE KARYAWAN SUDAH BENAR++++++++++++++++++++
    strsql = "SELECT 1 FROM h_attendance WHERE employee_code = '" & txtkdkar.Text & "' AND " _
        & "date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    If rs2.RecordCount = 0 Then
        MsgBox "Data Kehadiran Karyawan Ini Belum Diinput!" & Chr(13) & _
            "Silahkan Input Kehadiran Karyawan Terlebih Dahulu..", vbCritical, headerMSG
        rs2.Close
        Exit Sub
    End If
    rs2.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
'    '+++++++++++++++++++APAKAH KARYAWAN SUDAH PERNAH DI INPUT++++++++++++++++++++
'    If editTrans = False Then
'        strsql = "SELECT 1 FROM h_overtime WHERE employee_code = '" & txtkdkar.Text & "' " _
'            & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
'        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'        If rs2.RecordCount > 0 Then
'            MsgBox "This Employee Is Already Exist...!!", vbCritical
'            rs2.Close
'            Exit Sub
'        End If
'        rs2.Close
'    End If
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    Dim tglawal As Date, tglAkhir As Date
    Dim v_flag_active As Integer
    
'    strsql = "SELECT employee_code, flag_active FROM m_employee WHERE employee_code = '" & txtkdkar.Text & "'"
'    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'        v_flag_active = rs2!flag_active
'    rs2.Close
'
'    If v_flag_active = 2 Then
'        MsgBox "Employees in the MC condition, Update Attendance does not Allowed!"
'        Exit Sub
'    End If

    InsertPresent
    
    MsgBox "Save Succesfully..", vbInformation, headerMSG
    
    txtkdkar.Text = ""
    txt_nik.Text = ""
    txtnmkar.Text = ""
    txtkddiv.Text = ""
    txtdivision.Text = ""
    txtkdtitle.Text = ""
    txttitle.Text = ""
    txtjmljam.Text = 0
    txt15jam.Text = 0
    txt2jam.Text = 0
    Txt3jam.Text = 0
    Txt4jam.Text = 0
    txtket.Text = ""
    ttin.Value = "00:00"
    ttout.Value = "00:00"
    txt_nik.SetFocus
    editTrans = False
    
    frm_list_manual_overtime.isiGridAbsen
    'frm_list_manual_att.isiGridAbsen
End Sub

Private Sub vbbutton2_Click()
    isiGridKar (1)
End Sub

Private Sub vbButton5_Click()
    Unload Me
End Sub

Private Sub InsertPresent()
Dim start_time As String, end_time As String, max_break_out As String, min_break_in As String
Dim time_in As String, time_out_break As String, time_in_break As String, time_out As String
Dim tot_overtime, tot_ot As Double
Dim tt_in As String, tt_out As String
Dim rsot As New ADODB.Recordset

strsql = "SELECT total_ot, 15jam satujam, 2jam duajam, 3jam tigajam, 4jam empatjam, holiday, tot_overtime FROM h_attendance " _
        & "WHERE employee_code = '" & txtkdkar.Text & "' " _
        & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
rsot.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rsot.RecordCount > 0 Then
    v_tot_ot = rsot!total_OT
    v_15jam = rsot!satujam
    v_2jam = rsot!duajam
    v_3jam = rsot!tigajam
    v_4jam = rsot!empatjam
    v_holiday = rsot!holiday
    v_tot_overtime = rsot!tot_overtime
End If
rsot.Close

tot_overtime = ((1.5 * Val(txt15jam.Text)) + (2 * Val(txt2jam.Text)) + (3 * Val(Txt3jam.Text)) + (4 * Val(Txt4jam.Text)))

'tot_ot = ((1.5 * Val(v_15)) + (2 * Val(v_2)) + (3 * Val(v_3)) + (4 * Val(v_4)))

time_in = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(ttin.Value, "hh:mm") & ":00"
time_out = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    
tt_in = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(v_tt_in, "hh:mm") & ":00"
tt_out = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(v_tt_out, "hh:mm") & ":00"

    If editTrans = False Then
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        If v_ot_hour <> txtjmljam.Text Then
            strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_overtime.TDBCombo_company.Text & "'"
            CnG.Execute strsql
        End If
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
'        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
'        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_overtime.TDBCombo_company.Text & "'"
'        CnG.Execute strsql
'        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
        v_tot_ot = (Val(txtjmljam.Text) + v_tot_ot)

        Call totalOT
        tot_ot = ((1.5 * Val(v_15)) + (2 * Val(v_2)) + (3 * Val(v_3)) + (4 * Val(v_4)))
        
        strsql = "INSERT INTO h_overtime (employee_code,att_date,time_in,time_out,shift_code,total_hour,15jam,2jam," _
                & "3jam,4jam,ket,entry_date,entry_user,holiday,tot_overtime) VALUES " _
                & "('" & txtkdkar.Text & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & time_in & "','" & time_out & "','" & txtkdshift.Text & "'," _
                & "'" & Val(txtjmljam.Text) & "','" & Val(txt15jam.Text) & "','" & Val(txt2jam.Text) & "','" & Val(Txt3jam.Text) & "','" & Val(Txt4jam.Text) & "'," _
                & "'" & txtket.Text & "',now(),'" & LOGIN_CODE & "','" & Check1.Value & "','" & tot_overtime & "')"
        CnG.Execute strsql
         
        strsql = "UPDATE h_attendance set " _
                & "att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',shift_code = '" & txtkdshift.Text & "'," _
                & "total_ot = '" & v_tot_ot & "',15jam = '" & v_15 & "'," _
                & "2jam = '" & v_2 & "',3jam = '" & v_3 & "'," _
                & "4jam = '" & v_4 & "',holiday = '" & Check1.Value & "'," _
                & "tot_overtime = '" & tot_ot & "', editdate=now() " _
                & "WHERE employee_code = '" & txtkdkar.Text & "' " _
                & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
        CnG.Execute strsql
    Else
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        If v_ot_hour <> txtjmljam.Text Then
            strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_overtime.TDBCombo_company.Text & "'"
            CnG.Execute strsql
        End If
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
        Dim tot_hours As Double
        
        strsql = "SELECT total_hour FROM h_overtime " _
            & "WHERE employee_code = '" & txtkdkar.Text & "' " _
            & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " _
            & "AND time_in = '" & tt_in & "' " _
            & "AND time_out = '" & tt_out & "'"
        rsot.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsot.RecordCount > 0 Then
            tot_hours = rsot!total_hour
        End If
        rsot.Close
        
        v_tot_ot = (v_tot_ot - tot_hours + Val(txtjmljam.Text))

        Call totalOT
        tot_ot = ((1.5 * Val(v_15)) + (2 * Val(v_2)) + (3 * Val(v_3)) + (4 * Val(v_4)))
        
        strsql = "Update h_overtime set " _
                & "att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'," _
                & "time_in = '" & time_in & "', time_out = '" & time_out & "'," _
                & "shift_code = '" & txtkdshift.Text & "'," _
                & "total_hour = '" & txtjmljam.Text & "',15jam = '" & txt15jam.Text & "'," _
                & "2jam = '" & txt2jam.Text & "',3jam = '" & Txt3jam.Text & "'," _
                & "4jam = '" & Txt4jam.Text & "',holiday = '" & Check1.Value & "'," _
                & "tot_overtime = '" & tot_overtime & "',edit_date=now(),edit_user = '" & LOGIN_CODE & "' " _
                & "WHERE employee_code = '" & txtkdkar.Text & "' " _
                & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " _
                & "AND time_in = '" & tt_in & "' " _
                & "AND time_out = '" & tt_out & "'"
        CnG.Execute strsql
                
'        strsql = "UPDATE h_attendance set " _
'                & "att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',shift_code = '" & txtkdshift.Text & "'," _
'                & "total_ot = '" & txtjmljam.Text & "',15jam = '" & txt15jam.Text & "'," _
'                & "2jam = '" & txt2jam.Text & "',3jam = '" & Txt3jam.Text & "'," _
'                & "4jam = '" & Txt4jam.Text & "',holiday = '" & Check1.Value & "'," _
'                & "tot_overtime = '" & tot_overtime & "',editdate=now() " _
'                & "WHERE employee_code = '" & txtkdkar.Text & "' " _
'                & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
'        CnG.Execute strsql
        
        strsql = "UPDATE h_attendance set " _
                & "att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',shift_code = '" & txtkdshift.Text & "'," _
                & "total_ot = '" & v_tot_ot & "',15jam = '" & v_15 & "'," _
                & "2jam = '" & v_2 & "',3jam = '" & v_3 & "'," _
                & "4jam = '" & v_4 & "',holiday = '" & Check1.Value & "'," _
                & "tot_overtime = '" & tot_ot & "', editdate=now() " _
                & "WHERE employee_code = '" & txtkdkar.Text & "' " _
                & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
         CnG.Execute strsql
    End If

End Sub

Public Sub totalOT()
If Check1.Value = 1 Then
    v_tot_work = 0
    If v_tot_ot > 15 Then
        v_tot_ot = 15
    Else
        v_tot_ot = v_tot_ot
    End If
        Select Case Val(v_tot_ot)
            Case Is <= 7
                v_15 = 0
                v_2 = Val(v_tot_ot)
                v_3 = 0
                v_4 = 0
            Case Is = 7.5
                v_15 = 0
                v_2 = 7
                v_3 = 0.5
                v_4 = 0
            Case Is = 8
                v_15 = 0
                v_2 = 7
                v_3 = 1
                v_4 = 0
            Case Else
                v_15 = 0
                v_2 = 7
                v_3 = 1
                v_4 = Val(v_tot_ot) - Val(v_2) - Val(v_3)
        End Select
    Else
        Select Case Val(v_tot_ot)
            Case Is <= 1
                v_15 = 1
                v_2 = 0
                v_3 = 0
                v_4 = 0
            Case Is > 1
                v_15 = 1
                v_2 = Val(v_tot_ot) - Val(v_15)
                v_3 = 0
                v_4 = 0
            Case Else
                v_15 = 0
                v_2 = 0
                v_3 = 0
                v_4 = 0
        End Select
    End If
End Sub

Private Sub totalJam()
    Dim tglin As Date, tglout As Date
    Dim jambreak As Integer
    Dim tot_menit As Double
    Dim tgl_masuk As Date
    Dim tgl_keluar As Date
    Dim menit As Double
    Dim jam As Double
    Dim v_ot As Double
    
    Dim e As Long
    
    tglin = Date & " " & Format(ttin.Value, "hh:nn:ss")
    
    If Format(ttin.Value, "hh:nn:ss") > Format(ttout.Value, "hh:nn:s") Then
        tglout = Date + 1 & " " & ttout.Value
    Else
        tglout = Date & " " & ttout.Value
    End If
    
    tgl_masuk = Format(tglin, "yyyy-MM-dd hh:nn:ss")
    tgl_keluar = Format(tglout, "yyyy-MM-dd hh:nn:ss")
  
    menit = (DateDiff("n", tgl_masuk, tgl_keluar)) / 60
    jam = roundDown(menit)
    
    e = (menit - jam) * 60
    
    If e <= 20 Then
        tot_menit = 0
    ElseIf e > 20 And e <= 45 Then
        tot_menit = 0.5
    Else
        tot_menit = 1
    End If
    
txtjmljam.Text = jam + tot_menit

    If Check1.Value = 1 Then
        Select Case Val(txtjmljam.Text)
            Case Is <= 7
                txt15jam.Text = 0
                txt2jam.Text = Val(txtjmljam.Text)
                Txt3jam.Text = 0
                Txt4jam.Text = 0
            Case Is = 8
                txt15jam.Text = 0
                txt2jam.Text = 7
                Txt3jam.Text = 1
                Txt4jam.Text = 0
            Case Else
                txt15jam.Text = 0
                txt2jam.Text = 7
                Txt3jam.Text = 1
                Txt4jam.Text = Val(txtjmljam.Text) - Val(txt2jam.Text) - Val(Txt3jam.Text)
        End Select
    Else
        Select Case Val(txtjmljam.Text)
            Case Is <= 1
                txt15jam.Text = Val(txtjmljam.Text)
                txt2jam.Text = 0
                Txt3jam.Text = 0
                Txt4jam.Text = 0
            Case Is > 1
                txt15jam.Text = 1
                txt2jam.Text = Val(txtjmljam.Text) - Val(txt15jam.Text)
                Txt3jam.Text = 0
                Txt4jam.Text = 0
            Case Else
                txt15jam.Text = 0
                txt2jam.Text = 0
                Txt3jam.Text = 0
                Txt4jam.Text = 0
        End Select
    End If
End Sub

