VERSION 5.00
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_att_man 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MANUAL ATTENDANCE"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8490
   Icon            =   "frm_trans_att_man_new.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_shift_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3300
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   36
      Top             =   1170
      Width           =   3825
   End
   Begin prj_tpc.LynxGrid LynxGrid1 
      Height          =   3465
      Left            =   1650
      TabIndex        =   29
      Top             =   2580
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6112
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
   Begin VB.CheckBox chk_all_employee 
      Caption         =   "All Employee"
      Height          =   255
      Left            =   1650
      TabIndex        =   1
      Top             =   1980
      Width           =   1455
   End
   Begin VB.TextBox txt_department_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3300
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   33
      Top             =   1560
      Width           =   3825
   End
   Begin VB.TextBox txt_employee_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   3690
      TabIndex        =   16
      Top             =   2280
      Width           =   3435
   End
   Begin VB.TextBox txt_employee_code 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6750
      TabIndex        =   32
      Top             =   2280
      Visible         =   0   'False
      Width           =   825
   End
   Begin prj_tpc.vbButton cmdBrowse 
      Height          =   285
      Left            =   3300
      TabIndex        =   30
      Top             =   2280
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
      MICON           =   "frm_trans_att_man_new.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txt_status_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3300
      TabIndex        =   28
      Top             =   3360
      Width           =   3825
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1650
      TabIndex        =   26
      Top             =   780
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   155516931
      CurrentDate     =   40823
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   25
      Top             =   5970
      Width           =   7635
   End
   Begin VB.TextBox txt_title_code 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1650
      TabIndex        =   24
      Top             =   3000
      Width           =   1605
   End
   Begin VB.TextBox txt_division_code 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1650
      TabIndex        =   23
      Top             =   2640
      Width           =   1605
   End
   Begin VB.TextBox txt_nik 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1650
      TabIndex        =   2
      Top             =   2280
      Width           =   1605
   End
   Begin VB.TextBox txt_division_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   3300
      TabIndex        =   15
      Top             =   2640
      Width           =   3825
   End
   Begin VB.TextBox txt_title_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   3300
      TabIndex        =   14
      Top             =   3000
      Width           =   3825
   End
   Begin VB.TextBox txt_description 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1650
      MaxLength       =   50
      TabIndex        =   10
      Top             =   5610
      Width           =   5445
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_Status 
      Height          =   345
      Left            =   1650
      OleObjectBlob   =   "frm_trans_att_man_new.frx":05A6
      TabIndex        =   3
      Top             =   3360
      Width           =   1605
   End
   Begin prj_tpc.vbButton cmdSave 
      Height          =   705
      Left            =   510
      TabIndex        =   11
      Top             =   6300
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
      MICON           =   "frm_trans_att_man_new.frx":24FB
      PICN            =   "frm_trans_att_man_new.frx":2517
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
      Left            =   2160
      TabIndex        =   12
      Top             =   6300
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
      MICON           =   "frm_trans_att_man_new.frx":35A9
      PICN            =   "frm_trans_att_man_new.frx":35C5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_department 
      Height          =   375
      Left            =   1650
      OleObjectBlob   =   "frm_trans_att_man_new.frx":4657
      TabIndex        =   0
      Top             =   1560
      Width           =   1605
   End
   Begin TDBTime6Ctl.TDBTime ttout 
      Height          =   285
      Left            =   5430
      TabIndex        =   9
      Top             =   5250
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frm_trans_att_man_new.frx":65C0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frm_trans_att_man_new.frx":662C
      Spin            =   "frm_trans_att_man_new.frx":667C
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
   Begin TDBTime6Ctl.TDBTime ttin 
      Height          =   285
      Left            =   1650
      TabIndex        =   7
      Top             =   5250
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frm_trans_att_man_new.frx":66A4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frm_trans_att_man_new.frx":6710
      Spin            =   "frm_trans_att_man_new.frx":6760
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
   Begin VB.Frame Frame2 
      Caption         =   "Attendance Periode"
      Height          =   1425
      Left            =   1650
      TabIndex        =   34
      Top             =   3750
      Width           =   5445
      Begin VB.CheckBox chkLate 
         Caption         =   "Late"
         Height          =   225
         Left            =   1320
         TabIndex        =   42
         Top             =   1050
         Width           =   1575
      End
      Begin VB.CheckBox chkSun 
         Caption         =   "Sunday"
         Height          =   225
         Left            =   2460
         TabIndex        =   41
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkSat 
         Caption         =   "Saturday"
         Height          =   225
         Left            =   1320
         TabIndex        =   40
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frm_trans_att_man_new.frx":6788
         Left            =   2490
         List            =   "frm_trans_att_man_new.frx":678A
         TabIndex        =   5
         Text            =   "..."
         Top             =   210
         Width           =   645
      End
      Begin MSComCtl2.DTPicker DTPicker_From 
         Height          =   315
         Left            =   870
         TabIndex        =   4
         Top             =   210
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   155516931
         CurrentDate     =   40823
      End
      Begin MSComCtl2.DTPicker DTPicker_To 
         Height          =   315
         Left            =   3210
         TabIndex        =   6
         Top             =   210
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   155516931
         CurrentDate     =   40823
      End
      Begin VB.Label Label10 
         Caption         =   "NOT INCL."
         Height          =   195
         Left            =   300
         TabIndex        =   43
         Top             =   1050
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "NOT INCL."
         Height          =   195
         Left            =   300
         TabIndex        =   39
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label16 
         Caption         =   "FROM"
         Height          =   195
         Left            =   330
         TabIndex        =   35
         Top             =   270
         Width           =   495
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_shift 
      Height          =   375
      Left            =   1650
      OleObjectBlob   =   "frm_trans_att_man_new.frx":678C
      TabIndex        =   37
      Top             =   1170
      Width           =   1605
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SHIFT :"
      Height          =   195
      Left            =   360
      TabIndex        =   38
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT MANUAL ATTENDANCE"
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
      Left            =   390
      TabIndex        =   31
      Top             =   150
      Width           =   4845
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ABSENT STATUS :"
      Height          =   195
      Left            =   180
      TabIndex        =   27
      Top             =   3420
      Width           =   1425
   End
   Begin VB.Label Label5 
      Caption         =   "EMPLOYEE :"
      Height          =   210
      Left            =   630
      TabIndex        =   22
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "DIVISION :"
      Height          =   210
      Left            =   810
      TabIndex        =   21
      Top             =   2700
      Width           =   795
   End
   Begin VB.Label Label7 
      Caption         =   "TITLE :"
      Height          =   210
      Left            =   1050
      TabIndex        =   20
      Top             =   3060
      Width           =   555
   End
   Begin VB.Label Label9 
      Caption         =   "Remark :"
      Height          =   210
      Left            =   930
      TabIndex        =   18
      Top             =   5670
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DEPARTMENT :"
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "DATE :"
      Height          =   195
      Left            =   1050
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_att_man_new.frx":86E4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12540
   End
   Begin VB.Label Label8 
      Caption         =   "TIME IN :"
      Height          =   195
      Left            =   900
      TabIndex        =   19
      Top             =   5280
      Width           =   765
   End
   Begin VB.Label Label14 
      Caption         =   "TIME OUT :"
      Height          =   195
      Left            =   4440
      TabIndex        =   17
      Top             =   5280
      Width           =   945
   End
End
Attribute VB_Name = "frm_trans_att_man"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset
Dim rsStatus As New ADODB.Recordset
Dim rsShift As New ADODB.Recordset

Public editTrans As Boolean

'Dim tgl_masuk As Date
'Dim tgl_keluar As Date
Public v_tt_in, v_tt_out, v_absen_status As String
Dim tglAwal, tglAkhir As Date
Dim vParam As String

Public vGroupCode As String
Public vShiftCode As String

Private Sub createKar()
   With LynxGrid1
      .AddColumn "EMPLOYEE CODE", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "EMPLOYEE NAME", 2000, , , , , , , , , True
      .AddColumn "DIV. CODE", , , , , , , , , False
      .AddColumn "DIVISION", 2000, , , , , , , , , True
      .AddColumn "TITLE CODE", , , , , , , , , False
      .AddColumn "JOB TITLE", 2000, , , , , , , , , True
      .AddColumn "EMP. CODE", 2000, , , , , , , , False
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
Dim vFlagRollable As Integer

    If pilihan = 1 Then
        SQL = "SELECT flag_rollable FROM m_shift_group WHERE group_code = '" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            vFlagRollable = rscari!flag_rollable
        End If
        rscari.Close
    
        vParam = IIf(DEPARTMENT_CODE <> "" And DIVISION_CODE = "", "a.department_code = '" & DEPARTMENT_CODE & "'", IIf(DEPARTMENT_CODE = "" And DIVISION_CODE = "", "a.company_code = '" & COMPANY_CODE & "'", "a.department_code = '" & DEPARTMENT_CODE & "' AND a.division_code = '" & DIVISION_CODE & "'"))
        
        LynxGrid1.Clear
        If LOGIN_LEVEL = 100 Then
            If vFlagRollable = 0 Then
                If TDBCombo_department = "" Then
                    SQL = "select DISTINCT a.nik,a.employee_name,a.division_code, " & _
                            "b.division_name,a.title_code,c.title_name,a.employee_code " & _
                          "from m_employee a join m_division b on a.division_code = b.division_code and a.department_code = b.department_code and a.company_code = b.company_code " & _
                            "join m_title c on a.title_code = c.title_code " & _
                            "join td_shift2 d on a.employee_code = d.employee_code " & _
                          "WHERE a.flag_active <> 0 AND a.company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' " & _
                            "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                                "OR employee_name LIKE '%" & txt_nik.Text & "%')"
'                            "AND d.shift_code = '" & vShiftCode & "'"
                Else
                    SQL = "select DISTINCT a.nik,a.employee_name,a.division_code, " & _
                            "b.division_name,a.title_code,c.title_name,a.employee_code " & _
                         "from m_employee a join m_division b on a.division_code = b.division_code and a.department_code = b.department_code and a.company_code = b.company_code " & _
                            "join m_title c on a.title_code = c.title_code " & _
                            "join td_shift2 d on a.employee_code = d.employee_code " & _
                          "WHERE flag_active <> 0 AND a.company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' AND a.department_code = '" & TDBCombo_department.Text & "' " & _
                            "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                                "OR employee_name LIKE '%" & txt_nik.Text & "%')"
'                            "AND d.shift_code = '" & vShiftCode & "'"
                End If
            Else
                If TDBCombo_department = "" Then
                    SQL = "select DISTINCT a.nik,a.employee_name,a.division_code, " & _
                            "b.division_name,a.title_code,c.title_name,a.employee_code " & _
                          "from m_employee a join m_division b on a.division_code = b.division_code and a.department_code = b.department_code and a.company_code = b.company_code " & _
                            "join m_title c on a.title_code = c.title_code " & _
                            "join td_emp_group d on a.employee_code = d.employee_code " & _
                            "join tm_emp_group e on d.group_number = e.emp_group_number " & _
                            "join m_shift_new f on e.emp_group_code = f.group_code " & _
                          "WHERE a.flag_active <> 0 AND a.company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' " & _
                            "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                                "OR employee_name LIKE '%" & txt_nik.Text & "%')"
'                            "AND f.shift_code = '" & vShiftCode & "' AND date(f.shift_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
                Else
                    SQL = "select DISTINCT a.nik,a.employee_name,a.division_code, " & _
                            "b.division_name,a.title_code,c.title_name,a.employee_code " & _
                         "from m_employee a join m_division b on a.division_code = b.division_code and a.department_code = b.department_code and a.company_code = b.company_code " & _
                            "join m_title c on a.title_code = c.title_code " & _
                            "join td_emp_group d on a.employee_code = d.employee_code " & _
                            "join tm_emp_group e on d.group_number = e.emp_group_number " & _
                            "join m_shift_new f on e.emp_group_code = f.group_code " & _
                          "WHERE flag_active <> 0 AND a.company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' AND a.department_code = '" & TDBCombo_department.Text & "' " & _
                            "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                                "OR employee_name LIKE '%" & txt_nik.Text & "%')"
'                            "AND f.shift_code = '" & vShiftCode & "' AND date(f.shift_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
                End If
            End If
        Else
            If vFlagRollable = 0 Then
                SQL = "select DISTINCT a.nik,a.employee_name,a.division_code, " & _
                        "b.division_name,a.title_code,c.title_name,a.employee_code " & _
                      "from m_employee a join m_division b on a.division_code = b.division_code and a.department_code = b.department_code and a.company_code = b.company_code " & _
                        "join m_title c on a.title_code = c.title_code " & _
                        "join td_shift2 d on a.employee_code = d.employee_code " & _
                      "WHERE flag_active <> 0 AND a.company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' " & _
                        "AND " & vParam & " " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_nik.Text & "%') " & _
                        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) "
'                        "AND d.shift_code = '" & vShiftCode & "'"
            Else
                SQL = "select DISTINCT a.nik,a.employee_name,a.division_code, " & _
                          "b.division_name,a.title_code,c.title_name,a.employee_code " & _
                        "from m_employee a join m_division b on a.division_code = b.division_code and a.department_code = b.department_code and a.company_code = b.company_code " & _
                          "join m_title c on a.title_code = c.title_code " & _
                          "join td_emp_group d on a.employee_code = d.employee_code " & _
                            "join tm_emp_group e on d.group_number = e.emp_group_number " & _
                            "join m_shift_new f on e.emp_group_code = f.group_code " & _
                        "WHERE flag_active <> 0 AND a.company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' " & _
                          "AND " & vParam & " " & _
                          "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                              "OR employee_name LIKE '%" & txt_nik.Text & "%') " & _
                          "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) "
'                          "AND f.shift_code = '" & vShiftCode & "' AND date(f.shift_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
            End If
        End If
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid1.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid1.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME & vbTab & _
                rs!DIVISION_CODE & vbTab & rs!division_name & vbTab & _
                rs!title_code & vbTab & rs!title_name & vbTab & rs!employee_code
                rs.MoveNext
            Wend
            LynxGrid1.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                txt_employee_code.Text = rs!employee_code
                txt_nik.Text = rs!nik
                txt_employee_name.Text = rs!EMPLOYEE_NAME
                txt_division_code.Text = rs!DIVISION_CODE
                txt_division_name.Text = IIf(IsNull(rs!division_name) = True, "", rs!division_name)
                txt_title_code.Text = rs!title_code
                txt_title_name.Text = rs!title_name
            Else
                LynxGrid1.Visible = True
                LynxGrid1.SetFocus
            End If
        Else
            
        End If
        rs.Close
    Else
        If LynxGrid1.Rows > 0 Then
            txt_nik.Text = LynxGrid1.CellText(LynxGrid1.Row, 0)
            txt_employee_name.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
            txt_division_code.Text = LynxGrid1.CellText(LynxGrid1.Row, 2)
            txt_division_name.Text = LynxGrid1.CellText(LynxGrid1.Row, 3)
            txt_title_code.Text = LynxGrid1.CellText(LynxGrid1.Row, 4)
            txt_title_name.Text = LynxGrid1.CellText(LynxGrid1.Row, 5)
            txt_employee_code.Text = LynxGrid1.CellText(LynxGrid1.Row, 6)
        End If
        LynxGrid1.Visible = False
    End If
End Sub

Private Sub chk_all_employee_Click()
    If chk_all_employee.Value = 0 Then
        txt_nik.Enabled = True
        CmdBrowse.Enabled = True
    Else
        txt_nik.Enabled = False
        CmdBrowse.Enabled = False
    End If
End Sub

Private Sub Combo3_Click()
    If Combo3.Text = "TO" Then
        DTPicker_to.Value = DTPicker_from.Value
        DTPicker_to.Visible = True
    Else
        DTPicker_to.Visible = False
    End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmdSave_Click
    End If
End Sub

Private Sub Form_Load()
    Call createKar
    chk_all_employee.Value = 0
    
    Combo3.AddItem "..."
    Combo3.AddItem "TO"
    Combo3.Text = "..."
    
    If LOGIN_LEVEL = 100 Then
        Call load_data_department
    Else
        If DEPARTMENT_CODE <> "" Then
            SQL = "SELECT department_name FROM m_department " & _
                    "WHERE company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' " & _
                        "AND department_code = '" & DEPARTMENT_CODE & "'"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                TDBCombo_department.Text = DEPARTMENT_CODE
                txt_department_name.Text = rs!department_name
                TDBCombo_department.Enabled = False
                txt_department_name.Enabled = False
            End If
            rs.Close
        Else
            Call load_data_department
        End If
    End If
    Call load_data_status
    
    Call load_data_shift
End Sub

Private Sub LynxGrid1_DblClick()
    isiGridKar (2)
End Sub

Private Sub LynxGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        LynxGrid1.Visible = False
    End If
    If KeyAscii = 13 Then
        isiGridKar (2)
    End If
End Sub

Private Sub LynxGrid1_LostFocus()
    LynxGrid1.Visible = False
End Sub

Private Sub load_data_department()
    If rsDept.State Then rsDept.Close
    SQL = "select * from m_department " & _
            "where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' order by department_code"
    rsDept.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_department.RowSource = rsDept
End Sub

Private Sub load_data_status()
    If rsStatus.State Then rsStatus.Close
    SQL = "select * from m_absent_status WHERE kode < 7 order by kode"
    rsStatus.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_Status.RowSource = rsStatus
End Sub

Private Sub load_data_shift()
    If rsShift.State Then rsShift.Close
    SQL = "select * from m_shift where group_code = '" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "' order by shift_code"
    rsShift.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_shift.RowSource = rsShift
End Sub

Private Sub TDBCombo_department_ItemChange()
    If TDBCombo_department.ApproxCount > 0 Then
        TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
        txt_department_name.Text = TDBCombo_department.Columns("department_name").Value
    End If
End Sub

Public Sub TDBCombo_status_ItemChange()
    If TDBCombo_Status.ApproxCount > 0 Then
        TDBCombo_Status.Text = TDBCombo_Status.Columns("absent_code").Value
        txt_status_name.Text = TDBCombo_Status.Columns("absent_name").Value
    End If
    
    If TDBCombo_Status.Text <> "P" Then
        ttin.Enabled = False
        ttout.Enabled = False
    Else
        ttin.Enabled = True
        ttout.Enabled = True
    End If
End Sub

Private Sub TDBCombo_shift_ItemChange()
    If TDBCombo_shift.ApproxCount > 0 Then
        TDBCombo_shift.Text = TDBCombo_shift.Columns("shift_code").Value
        txt_shift_name.Text = TDBCombo_shift.Columns("shift_name").Value
    End If
End Sub

Private Sub ttout_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_description.SetFocus
    End If
End Sub

Private Sub txt_nik_Change()
    If txt_nik.Text = "" Then
        txt_employee_name.Text = ""
        txt_division_code.Text = ""
        txt_division_name.Text = ""
        txt_title_code.Text = ""
        txt_title_name.Text = ""
    End If
End Sub

Private Sub txt_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub txt_description_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSave.SetFocus
    End If
End Sub

Private Sub cmdSave_Click()
Dim rsemp As New ADODB.Recordset
Dim aa As Integer
Dim a, b, c, d As String

'On Error GoTo Err
    If TDBCombo_Status.Text = "P" Then
        If ttin.Value = ttout.Value Then
            MsgBox "Invalid Check In Time...!!!", vbExclamation, headerMSG
            Exit Sub
        End If
        
        If Format(ttout.Value, "hh:mm") = "" Or Format(ttout.Value, "hh:mm") = "" Then
            MsgBox "Invalid Format Check In or Check Out Time....!!!", vbExclamation, headerMSG
            Exit Sub
        End If
    End If
    
    
    '+++++++++++++++++++APAKAH KODE KARYAWAN SUDAH BENAR++++++++++++++++++++
    If chk_all_employee.Value = 0 Then
        SQL = "SELECT employee_code FROM m_employee WHERE employee_code = '" & txt_employee_code.Text & "'"
        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount = 0 Then
            MsgBox "Invalid Employee Code...!!", vbCritical, headerMSG
            rs2.Close
            Exit Sub
        End If
        rs2.Close
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    tglAwal = DTPicker_from.Value
    If Combo3.Text = "TO" Then
        tglAkhir = DTPicker_to.Value
    Else
        tglAkhir = DTPicker_from.Value
    End If

    '+++++++++++++++++++APAKAH KARYAWAN SUDAH PERNAH DI INPUT++++++++++++++++++++
    If editTrans = False Then
        SQL = "SELECT att_date,shift_code,shift_number,status FROM h_attendance WHERE employee_code = '" & txt_employee_code.Text & "' " _
            & "AND att_date = '" & Format(tglAwal, "yyyy-MM-dd HH:mm:ss") & "'"
        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
        If rs2.RecordCount > 0 Then
            Dim v_shift As String
            v_shift = IIf(rs2!Status = "P", "PRESENT", IIf(rs2!Status = "A", "ALPHA", IIf(rs2!Status = "DT", "DUTY", _
                    IIf(rs2!Status = "OF", "OFF", IIf(rs2!Status = "PR", "PERMISSION", "Sakit Ada Surat Dokter")))))
            
            MsgBox "This Employee Is Already Exist With Status " & v_shift, vbCritical
            rs2.Close
            Exit Sub
        End If
        rs2.Close
    End If
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    Select Case TDBCombo_Status.Text
        Case "P", "OT"
            CnG.BeginTrans
            While tglAwal <= tglAkhir
                Dim vDayName As String
                Dim vHoliday As Integer
                Dim rsHol As New ADODB.Recordset
                
                SQL = "SELECT * FROM t_holiday WHERE date(holiday_date) = '" & Format(tglAwal, "yyyy-MM-dd") & "'"
                rsHol.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rsHol.RecordCount > 0 Then
                    vHoliday = 1
                Else
                    vHoliday = 0
                End If
                rsHol.Close
                
                vDayName = Format(tglAwal, "dddd")
                If chkSat.Value = 1 And chkSun.Value = 1 Then
                    If vHoliday = 0 Then
                        If vDayName <> "Saturday" Then
                            If vDayName <> "Sunday" Then
                                InsertPresent
                            End If
                        End If
                    End If
                ElseIf chkSat.Value = 1 And chkSun.Value = 0 Then
                    If vHoliday = 0 Then
                        If vDayName <> "Saturday" Then InsertPresent
                    End If
                ElseIf chkSat.Value = 0 And chkSun.Value = 1 Then
                    If vHoliday = 0 Then
                        If vDayName <> "Sunday" Then InsertPresent
                    End If
                ElseIf chkSat.Value = 0 And chkSun.Value = 0 Then
                    If vHoliday = 0 Then
                        InsertPresent
                    End If
                End If
                
                If editTrans = True Then
                    InsertPresent
                End If
                
                If tglAwal = tglAkhir Then
                    InsertPresent
                End If
                
                tglAwal = tglAwal + 1
            Wend
            CnG.CommitTrans
        Case "A", "OF", "PR", "S"
            Dim abstatus As Integer
            If TDBCombo_Status.Text = "A" Then
                abstatus = 2
            ElseIf TDBCombo_Status.Text = "OF" Then
                abstatus = 6
            ElseIf TDBCombo_Status.Text = "PR" Then
                abstatus = 0
            ElseIf TDBCombo_Status.Text = "S" Then
                abstatus = 1
            End If
            
            CnG.BeginTrans
                While tglAwal <= tglAkhir
                    If chk_all_employee.Value = 1 Then
                        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
                        SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                        CnG.Execute SQL
                        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        
                        If TDBCombo_department.Text = "" Then
                            SQL = "SELECT employee_code from m_employee " _
                                    & "where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                        Else
                            SQL = "SELECT employee_code from m_employee " _
                                    & "where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' AND " _
                                    & "department_code = '" & TDBCombo_department.Text & "'"
                        End If
                        rsemp.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                        
                        For aa = 0 To rsemp.RecordCount - 1
                            SQL = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "' AND employee_code = '" & rsemp!employee_code & "'"
                            CnG.Execute SQL
                            
                            SQL = "INSERT INTO h_attendance (att_date,employee_code,group_code,shift_code,status,flag_present,absent_status," & _
                                    "description,entry_date,userinput,shift_number,flag_manual) " & _
                                  "VALUES (" & _
                                    "'" & Format(tglAwal, "yyyy-MM-dd") & "','" & rsemp!employee_code & "','" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "'," & _
                                    "'" & frm_list_manual_att.TDBGrid_Shift.Columns("shift_code").Value & "','" & TDBCombo_Status.Text & "',0,'" & abstatus & "'," & _
                                    "'" & txt_description.Text & "',now(),'" & LOGIN_NAME & "',1,1)"
                            CnG.Execute SQL
                            rsemp.MoveNext
                        Next
                        rsemp.Close
                    Else
                        '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
                        If v_absen_status <> TDBCombo_Status.Text Then
                            SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                            CnG.Execute SQL
                        End If
                        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        
                        SQL = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "' AND employee_code = '" & txt_employee_code.Text & "'"
                        CnG.Execute SQL
                        
                        SQL = "INSERT INTO h_attendance (att_date,employee_code,group_code,shift_code,status,flag_present,absent_status," & _
                                "description,entry_date,userinput,shift_number,flag_manual) " & _
                              "VALUES (" & _
                                "'" & Format(tglAwal, "yyyy-MM-dd") & "','" & txt_employee_code.Text & "','" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "'," & _
                                "'" & frm_list_manual_att.TDBGrid_Shift.Columns("shift_code").Value & "','" & TDBCombo_Status.Text & "',0,'" & abstatus & "'," & _
                                "'" & txt_description.Text & "',now(),'" & LOGIN_NAME & "',1,1)"
                        CnG.Execute SQL
                    End If
                    tglAwal = tglAwal + 1

'                    CnG.Execute "call spg_leave_periode2 ('" & Format(tglAwal, "yyyy-MM-dd") & "')"
                
                Wend
            CnG.CommitTrans
        Case "DT"
            Dim waktu_in As String, waktu_out As String, waktu_break_out As String, waktu_break_in As String
            
            CnG.BeginTrans
                
            While tglAwal <= tglAkhir
                SQL = "SELECT start_time, end_time FROM m_shift " & _
                      "WHERE group_code = '" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "' " & _
                        "AND shift_code = '" & frm_list_manual_att.TDBGrid_Shift.Columns("shift_code").Value & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                    waktu_in = Format(tglAwal, "yyyy-MM-dd") & " " & Format(rs!start_time, "hh:mm") & ":00"
                    waktu_out = Format(tglAwal, "yyyy-MM-dd") & " " & Format(rs!end_time, "hh:mm") & ":00"
                Else
                    waktu_in = Format(tglAwal, "yyyy-MM-dd") & " 07:00:00"
                    waktu_out = Format(tglAwal, "yyyy-MM-dd") & " 15:00:00"
                End If
                rs.Close
                            
                If chk_all_employee.Value = 1 Then
                    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
                    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                    CnG.Execute SQL
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                    If TDBCombo_department.Text = "" Then
                        SQL = "SELECT employee_code from m_employee " _
                                & "where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                    Else
                        SQL = "SELECT employee_code from m_employee " _
                                & "where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' AND " _
                                & "department_code = '" & TDBCombo_department.Text & "'"
                    End If
                    rsemp.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    For aa = 0 To rsemp.RecordCount - 1
                        SQL = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "' AND employee_code = '" & rsemp!employee_code & "'"
                        CnG.Execute SQL
                            
                        SQL = "INSERT INTO h_attendance (att_date,employee_code,group_code,shift_code,status,time_in,time_out," & _
                                "flag_present,flag_duty,absent_status,description,entry_date,userinput,shift_number,flag_manual) " & _
                              "VALUES (" & _
                                "'" & Format(tglAwal, "yyyy-MM-dd") & "','" & rsemp!employee_code & "','" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "'," & _
                                "'" & frm_list_manual_att.TDBGrid_Shift.Columns("shift_code").Value & "','" & TDBCombo_Status.Text & "'," & _
                                "'" & waktu_in & "','" & waktu_out & "'," & _
                                "1,1,0,'" & txt_description.Text & "',now(),'" & LOGIN_NAME & "',1,1)"
                        CnG.Execute SQL
                        rsemp.MoveNext
                    Next
                    rsemp.Close
                Else
                    '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
                    If v_absen_status <> TDBCombo_Status.Text Then
                        SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                        CnG.Execute SQL
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                    SQL = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "' AND employee_code = '" & txt_employee_code.Text & "'"
                    CnG.Execute SQL
                        
                    SQL = "INSERT INTO h_attendance (att_date,employee_code,group_code,shift_code,status,time_in,time_out," & _
                            "flag_present,flag_duty,absent_status,description,entry_date,userinput,shift_number,flag_manual) " & _
                          "VALUES (" & _
                            "'" & Format(tglAwal, "yyyy-MM-dd") & "','" & txt_employee_code.Text & "','" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "'," & _
                            "'" & frm_list_manual_att.TDBGrid_Shift.Columns("shift_code").Value & "','" & TDBCombo_Status.Text & "'," & _
                            "'" & waktu_in & "','" & waktu_out & "'," & _
                            "1,1,0,'" & txt_description.Text & "',now(),'" & LOGIN_NAME & "',1,1)"
                    CnG.Execute SQL
                End If
                tglAwal = tglAwal + 1
                
'                CnG.Execute "call spg_leave_periode2 ('" & Format(tglAwal, "yyyy-MM-dd") & "')"
            
            Wend
            CnG.CommitTrans
        Case Else
            MsgBox "Please Check Absent Status....!!!", vbExclamation, headerMSG
            Exit Sub
    End Select
    
    MsgBox "Save Data Succesfully!", vbInformation, "Sukses!"
    
    Call clear_view_data
    
    If editTrans = True Then
        Unload Me
    Else
        editTrans = False
    End If
    
    frm_list_manual_att.load_data_att
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdBrowse_Click()
    isiGridKar (1)
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub InsertPresent()
Dim start_time As String, end_time As String, max_break_out As String, min_break_in As String
Dim time_in As String, time_out_break As String, time_in_break As String, time_out As String
Dim rsemp As New ADODB.Recordset
Dim vFlagRollable As Integer
   
    '+++++++++++++++++++++MENCARI TANGGAL BATAS JAM MASUK,ISTIRAHAT,KELUAR++++++++
    SQL = "Select CAST(concat('" & Format(tglAwal, "yyyy-MM-dd") & "',' ', time(start_time)) as datetime) start_time," _
            & "CAST(concat('" & Format(tglAwal, "yyyy-MM-dd") & "',' ', time(end_time)) as datetime) end_time," _
            & "CAST(concat('" & Format(tglAwal, "yyyy-MM-dd") & "',' ', time(max_break_out)) as datetime) max_break_out," _
            & "CAST(concat('" & Format(tglAwal, "yyyy-MM-dd") & "',' ', time(min_break_in)) as datetime) min_break_in," _
            & "curdate() tglserver " _
            & "from m_shift where shift_code = '" & frm_list_manual_att.TDBGrid_Shift.Columns("shift_code").Value & "' " & _
                "AND group_code = '" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "'"
'        & "from m_shift where shift_code = '" & txtkdshift.Text & "'"
    rs2.Open SQL, CnG, adOpenDynamic, adLockReadOnly
    
    start_time = Format(rs2!start_time, "yyyy-MM-dd hh:mm:ss")
    If frm_list_manual_att.TDBGrid_Shift.Columns("flag_day_over").Value = 1 Then
        end_time = Format(DateAdd("d", 1, rs2!end_time), "yyyy-MM-dd hh:mm:ss")
    Else
        end_time = Format(rs2!end_time, "yyyy-MM-dd hh:mm:ss")
    End If
    
    time_in = Format(tglAwal, "yyyy-MM-dd") & " " & Format(ttin.Value, "hh:mm") & ":00"
    
    If Format(ttout.Value, "hh:mm:ss") < Format(ttin.Value, "hh:mm:ss") Then
        time_out = Format(tglAwal + 1, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    Else
        time_out = Format(tglAwal, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    End If
    
    rs2.Close
    
    If chk_all_employee.Value = 1 Then
        If rscari.State Then rscari.Close
        SQL = "SELECT flag_rollable FROM m_shift_group WHERE group_code = '" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            vFlagRollable = IIf(IsNull(rscari!flag_rollable), 0, rscari!flag_rollable)
        Else
            vFlagRollable = 0
        End If
        rscari.Close
        
        If vFlagRollable = 0 Then
            SQL = "SELECT DISTINCT a.employee_code " & _
                  "FROM m_employee a JOIN td_shift b ON a.employee_code = b.employee_code " & _
                    "JOIN tm_shift c ON b.shift_number = c.shift_number AND b.group_code = c.group_code " & _
                    "JOIN m_shift_group d ON b.group_code = d.group_code " & _
                  "WHERE b.group_code = '" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "' "
            
            If TDBCombo_department.Text <> "" Then
                SQL = SQL & _
                        "AND a.department_code = '" & TDBCombo_department.Text & "'"
            End If
        Else
            SQL = "SELECT DISTINCT a.employee_code " & _
                  "FROM m_employee a JOIN td_emp_group b ON a.employee_code = b.employee_code " & _
                    "JOIN tm_emp_group c ON b.group_number = c.emp_group_number AND b.emp_group_code = c.emp_group_code " & _
                    "JOIN m_shift_group d ON b.group_code = d.group_code " & _
                  "WHERE b.group_code = '" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "' "
            
            If TDBCombo_department.Text <> "" Then
                SQL = SQL & _
                        "AND a.department_code = '" & TDBCombo_department.Text & "'"
            End If
        End If
    
        rsemp.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        For aa = 0 To rsemp.RecordCount - 1
            
            If editTrans = False Then
                '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
                SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                CnG.Execute SQL
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                
                SQL = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglAwal, "yyyy-MM-dd 00:00:00") & "' AND employee_code = '" & rsemp!employee_code & "'"
                CnG.Execute SQL
                
                SQL = "INSERT INTO h_attendance (employee_code,att_date,group_code,shift_code,status," & _
                        "shift_number,start_time,end_time,time_in,time_out,description,entry_date,userinput," & _
                        "absent_status,flag_present,flag_duty,flag_inc_late, flag_manual) " & _
                      "VALUES (" & _
                        "'" & rsemp!employee_code & "','" & Format(tglAwal, "yyyy-MM-dd") & "','" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "'," & _
                        "'" & TDBCombo_shift.Text & "','" & TDBCombo_Status.Text & "'," & _
                        "1,'" & start_time & "','" & end_time & "','" & time_in & "','" & time_out & "'," & _
                        "'" & txt_description.Text & "',now(),'" & LOGIN_NAME & "',0,1,0," & chkLate.Value & ",1)"
            Else
                '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
                If v_absen_status <> TDBCombo_Status.Text Then
                    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                    CnG.Execute SQL
                ElseIf v_tt_in <> ttin.Value Or v_tt_out <> ttout.Value Then
                    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                    CnG.Execute SQL
                End If
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                SQL = "UPDATE h_attendance set employee_code = '" & rsemp!employee_code & "'," & _
                        "att_date = '" & Format(tglAwal, "yyyy-MM-dd") & "',group_code = '" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "'," & _
                        "shift_code = '" & TDBCombo_shift.Text & "'," & _
                        "status = '" & TDBCombo_Status.Text & "'," & _
                        "shift_number = 1,start_time = '" & start_time & "',end_time = '" & end_time & "'," & _
                        "time_in = '" & time_in & "',flag_present = 1," & _
                        "time_out = '" & time_out & "'," & _
                        "description = '" & txt_description.Text & "',edit_date = now(),useredit = '" & LOGIN_NAME & "'," & _
                        "absent_status = 0, flag_present = 1, flag_duty = 0,flag_inc_late = " & chkLate.Value & " " & _
                        "WHERE employee_code = '" & rsemp!employee_code & "' " & _
                            "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") & "'"
        '                & "AND shift_code = '" & txtkdshift.Text & "'"
            End If
            CnG.Execute SQL
            rsemp.MoveNext
        Next
    Else
        If editTrans = False Then
            '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
            SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
            CnG.Execute SQL
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            SQL = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglAwal, "yyyy-MM-dd 00:00:00") & "' AND employee_code = '" & txt_employee_code.Text & "'"
            CnG.Execute SQL
                
            SQL = "INSERT INTO h_attendance (employee_code,att_date,group_code,shift_code,status," & _
                    "shift_number,start_time,end_time,time_in,time_out,description,entry_date,userinput," & _
                    "absent_status,flag_present,flag_duty,flag_inc_late, flag_manual) " & _
                  "VALUES (" & _
                    "'" & txt_employee_code.Text & "','" & Format(tglAwal, "yyyy-MM-dd") & "','" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "'," & _
                    "'" & TDBCombo_shift.Text & "','" & TDBCombo_Status.Text & "'," & _
                    "1,'" & start_time & "','" & end_time & "','" & time_in & "','" & time_out & "'," & _
                    "'" & txt_description.Text & "',now(),'" & LOGIN_NAME & "',0,1,0," & chkLate.Value & ",1)"
        Else
            '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
            If v_absen_status <> TDBCombo_Status.Text Then
                SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                CnG.Execute SQL
            ElseIf v_tt_in <> ttin.Value Or v_tt_out <> ttout.Value Then
                SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                CnG.Execute SQL
            End If
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                
            SQL = "UPDATE h_attendance set employee_code = '" & txt_employee_code.Text & "'," & _
                    "att_date = '" & Format(tglAwal, "yyyy-MM-dd") & "',group_code = '" & frm_list_manual_att.TDBCombo_Group_Shift.Text & "'," & _
                    "shift_code = '" & TDBCombo_shift.Text & "'," & _
                    "status = '" & TDBCombo_Status.Text & "'," & _
                    "shift_number = 1,start_time = '" & start_time & "',end_time = '" & end_time & "'," & _
                    "time_in = '" & time_in & "',flag_present = 1," & _
                    "time_out = '" & time_out & "'," & _
                    "description = '" & txt_description.Text & "',edit_date = now(),useredit = '" & LOGIN_NAME & "'," & _
                    "absent_status = 0, flag_present = 1, flag_duty = 0,flag_inc_late =" & chkLate.Value & " " & _
                    "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                        "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") & "'"
    '                & "AND shift_code = '" & txtkdshift.Text & "'"
        End If
        CnG.Execute SQL
    End If
    
'    CnG.Execute "call spg_leave_periode2 ('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "')"

End Sub

Private Function check_exist_new_attendance(empCode As String, attDate As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
    check_exist_new_attendance = False

    str_sql = "select count(employee_code and att_date) as rec_count from h_attendance where employee_code = '" & _
                Replace$(Trim$(empCode), Chr$(39), Chr$(96)) & "' AND att_date = '" & attDate & "' "
    rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_exist_new_attendance = True
        Exit Function
    End If
End Function

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_company" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is DTPicker Then
            Ctr.Value = Now
        ElseIf TypeOf Ctr Is TDBTime Then
            Ctr.Value = "00:00"
        End If
    Next
End Sub


