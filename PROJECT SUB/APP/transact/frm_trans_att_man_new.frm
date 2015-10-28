VERSION 5.00
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_att_man 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MANUAL ATTENDANCE"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8490
   Icon            =   "frm_trans_att_man_new.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prj_absensi.LynxGrid LynxGrid1 
      Height          =   3465
      Left            =   1650
      TabIndex        =   18
      Top             =   2160
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
   Begin VB.TextBox txt_shift_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3300
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   21
      Top             =   1170
      Width           =   3825
   End
   Begin VB.CheckBox chk_all_employee 
      Caption         =   "All Employee"
      Height          =   255
      Left            =   1650
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txt_employee_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   3690
      TabIndex        =   9
      Top             =   1860
      Width           =   3435
   End
   Begin prj_absensi.vbButton cmdBrowse 
      Height          =   285
      Left            =   3300
      TabIndex        =   19
      Top             =   1860
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1650
      TabIndex        =   17
      Top             =   780
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   100204547
      CurrentDate     =   40823
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   7635
   End
   Begin VB.TextBox txt_title_code 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1650
      TabIndex        =   15
      Top             =   2220
      Width           =   1605
   End
   Begin VB.TextBox txt_employee_code 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1650
      TabIndex        =   1
      Top             =   1860
      Width           =   1605
   End
   Begin VB.TextBox txt_title_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   3300
      TabIndex        =   8
      Top             =   2220
      Width           =   3825
   End
   Begin VB.TextBox txt_description 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1650
      MaxLength       =   50
      TabIndex        =   5
      Top             =   4140
      Width           =   5445
   End
   Begin prj_absensi.vbButton cmdSave 
      Height          =   705
      Left            =   540
      TabIndex        =   6
      Top             =   5010
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
      MICON           =   "frm_trans_att_man_new.frx":05A6
      PICN            =   "frm_trans_att_man_new.frx":05C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_absensi.vbButton cmdExit 
      Height          =   705
      Left            =   2160
      TabIndex        =   7
      Top             =   5010
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
      MICON           =   "frm_trans_att_man_new.frx":1654
      PICN            =   "frm_trans_att_man_new.frx":1670
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBTime6Ctl.TDBTime ttout 
      Height          =   285
      Left            =   5430
      TabIndex        =   4
      Top             =   3780
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frm_trans_att_man_new.frx":2702
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frm_trans_att_man_new.frx":276E
      Spin            =   "frm_trans_att_man_new.frx":27BE
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
      TabIndex        =   2
      Top             =   3780
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frm_trans_att_man_new.frx":27E6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frm_trans_att_man_new.frx":2852
      Spin            =   "frm_trans_att_man_new.frx":28A2
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
      Height          =   1125
      Left            =   1650
      TabIndex        =   23
      Top             =   2580
      Width           =   5445
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frm_trans_att_man_new.frx":28CA
         Left            =   2490
         List            =   "frm_trans_att_man_new.frx":28CC
         TabIndex        =   26
         Text            =   "..."
         Top             =   210
         Width           =   645
      End
      Begin VB.CheckBox chkSun 
         Caption         =   "Sunday"
         Height          =   225
         Left            =   2460
         TabIndex        =   24
         Top             =   720
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker_From 
         Height          =   315
         Left            =   870
         TabIndex        =   27
         Top             =   210
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   100204547
         CurrentDate     =   40823
      End
      Begin MSComCtl2.DTPicker DTPicker_To 
         Height          =   315
         Left            =   3210
         TabIndex        =   28
         Top             =   210
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   100204547
         CurrentDate     =   40823
      End
      Begin VB.CheckBox chkSat 
         Caption         =   "Saturday"
         Height          =   225
         Left            =   1320
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "FROM"
         Height          =   195
         Left            =   330
         TabIndex        =   30
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "NOT INCL."
         Height          =   195
         Left            =   300
         TabIndex        =   29
         Top             =   720
         Width           =   795
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_shift 
      Height          =   375
      Left            =   1650
      OleObjectBlob   =   "frm_trans_att_man_new.frx":28CE
      TabIndex        =   31
      Top             =   1170
      Width           =   1575
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SHIFT :"
      Height          =   195
      Left            =   360
      TabIndex        =   22
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
      TabIndex        =   20
      Top             =   150
      Width           =   4845
   End
   Begin VB.Label Label5 
      Caption         =   "EMPLOYEE :"
      Height          =   210
      Left            =   630
      TabIndex        =   14
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "TITLE :"
      Height          =   210
      Left            =   1050
      TabIndex        =   13
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label Label9 
      Caption         =   "Remark :"
      Height          =   210
      Left            =   930
      TabIndex        =   11
      Top             =   4200
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "DATE :"
      Height          =   195
      Left            =   1050
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_att_man_new.frx":4CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12540
   End
   Begin VB.Label Label8 
      Caption         =   "TIME IN :"
      Height          =   195
      Left            =   900
      TabIndex        =   12
      Top             =   3810
      Width           =   765
   End
   Begin VB.Label Label14 
      Caption         =   "TIME OUT :"
      Height          =   195
      Left            =   4440
      TabIndex        =   10
      Top             =   3810
      Width           =   945
   End
End
Attribute VB_Name = "frm_trans_att_man"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim rsShift As New ADODB.Recordset

Public editTrans As Boolean

Public v_tt_in, v_tt_out, v_absen_status, vShiftCode As String
Dim tglAwal, tglAkhir As Date
Dim vParam As String

Dim vEnrollNumber As Double

Dim i As Integer
Dim m As Integer

Private Sub createKar()
    LynxGrid1.ClearAll
    
    With LynxGrid1
       .AddColumn "EMP. CODE", 1500, lgAlignCenterCenter, , , , , , , True
       .AddColumn "EMP. NAME", 2000, , , , , , , , , True
       .AddColumn "TITLE CODE", , , , , , , , , False
       .AddColumn "JOB TITLE", 2000, , , , , , , , , True
       .BackColorBkg = &HFCE1CB
       .Redraw = True
    End With
End Sub

Private Sub isiGridKar(pilihan As Integer)
Dim vFlagRollable As Integer

    vParam = IIf(LOGIN_LEVEL <> 100, _
                IIf(COMPANY_ACCESS = 0, _
                    "a.company_code = '" & COMPANY_CODE & "' ", _
                            "a.company_code = '" & frm_list_manual_att.TDBCombo_company().Columns("company_code").Value & "' "), "a.company_code = '" & frm_list_manual_att.TDBCombo_company().Columns("company_code").Value & "' ")
                        
    If pilihan = 1 Then
        LynxGrid1.Clear
        If rs2.State Then rs2.Close
        SQL = "select a.employee_code,a.employee_name, " & _
                "a.title_code,b.title_name " & _
             "from m_employee a join m_title b on a.title_code = b.title_code " & _
              "WHERE flag_active <> 0 " & _
                    "AND a.company_code = '" & frm_list_manual_att.TDBCombo_company().Columns("company_code").Value & "' " & _
                    "AND (employee_code LIKE '%" & txt_employee_code.Text & "%' " & _
                        "OR employee_name LIKE '%" & txt_employee_code.Text & "%')"
        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs2.RecordCount > 0 Then
            LynxGrid1.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid1.AddItem rs2!employee_code & vbTab & rs2!EMPLOYEE_NAME & vbTab & _
                rs2!title_code & vbTab & rs2!title_name
                rs2.MoveNext
            Wend
            LynxGrid1.Redraw = True
            If rs2.RecordCount = 1 Then
                rs2.MoveFirst
                txt_employee_code.Text = rs2!employee_code
                txt_employee_name.Text = rs2!EMPLOYEE_NAME
                txt_title_code.Text = rs2!title_code
                txt_title_name.Text = rs2!title_name
            Else
                LynxGrid1.Visible = True
                LynxGrid1.SetFocus
            End If
        Else
            
        End If
        rs2.Close
    Else
        If LynxGrid1.Rows > 0 Then
            txt_employee_code.Text = LynxGrid1.CellText(LynxGrid1.Row, 0)
            txt_employee_name.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
            txt_title_code.Text = LynxGrid1.CellText(LynxGrid1.Row, 4)
            txt_title_name.Text = LynxGrid1.CellText(LynxGrid1.Row, 5)
        End If
        LynxGrid1.Visible = False
    End If
End Sub

Private Sub chk_all_employee_Click()
    If chk_all_employee.Value = 0 Then
        txt_employee_code.Enabled = True
        cmdBrowse.Enabled = True
    Else
        txt_employee_code.Enabled = False
        cmdBrowse.Enabled = False
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

Public Sub set_data_shift(ByVal str_code As String)
On Error GoTo Err
    
    If rsShift.RecordCount > 0 Then
        rsShift.MoveFirst
        rsShift.Find ("shift_code='" & str_code & "'")   ', 0, adSearchForward, 1)
        If Not (rsShift.EOF = True Or rsShift.BOF = True) Then
            TDBCombo_shift.Bookmark = rsShift.AbsolutePosition
            Call TDBCombo_shift_ItemChange
        Else
            TDBCombo_shift.Text = ""
        End If
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub load_data_shift()
    If rsShift.State Then rsShift.Close
    SQL = "select * from m_shift where flag_shift = 0 order by shift_code"
    rsShift.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_shift.RowSource = rsShift
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

Private Sub txt_employee_code_Change()
    If txt_employee_code.Text = "" Then
        txt_employee_name.Text = ""
        txt_title_code.Text = ""
        txt_title_name.Text = ""
    End If
End Sub

Private Sub txt_employee_code_KeyPress(KeyAscii As Integer)
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
    If ttin.Value = ttout.Value Then
        MsgBox "Invalid Check In Time...!!!", vbExclamation, headerMSG
        Exit Sub
    End If
    
    If Format(ttout.Value, "hh:mm") = "" Or Format(ttout.Value, "hh:mm") = "" Then
        MsgBox "Invalid Format Check In or Check Out Time....!!!", vbExclamation, headerMSG
        Exit Sub
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
        If rs2.State Then rs2.Close
        SQL = "SELECT att_date,shift_code,shift_number FROM h_attendance WHERE employee_code = '" & txt_employee_code.Text & "' " _
            & "AND DATE(att_date) = '" & Format(tglAwal, "yyyy-MM-dd") & "'"
        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
        If rs2.RecordCount > 0 Then
            MsgBox "This employee already exist on attendance...", vbExclamation, headerMSG
            rs2.Close
            Exit Sub
        End If
        rs2.Close
    End If
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    CnG.BeginTrans
    While tglAwal <= tglAkhir
        m = m + 1
        
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
            InsertPresent
        End If
    
        tglAwal = tglAwal + 1
    Wend
    CnG.CommitTrans

    MsgBox "Save successfully...", vbInformation, headerMSG
    
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
Dim vShiftNumber As Integer

    '+++++++++++++++++++++MENCARI TANGGAL BATAS JAM MASUK,ISTIRAHAT,KELUAR++++++++
    If rs2.State Then rs2.Close
    SQL = "Select CAST(concat('" & Format(tglAwal, "yyyy-MM-dd") & "',' ', time(start_time)) as datetime) start_time," _
            & "CAST(concat('" & Format(tglAwal, "yyyy-MM-dd") & "',' ', time(end_time)) as datetime) end_time," _
            & "curdate() tglserver " _
            & "from m_shift where shift_code = '" & TDBCombo_shift.Columns("shift_code").Value & "'"
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
        If rsemp.State Then rsemp.Close
        SQL = "SELECT DISTINCT a.employee_code " & _
              "FROM m_employee a JOIN td_shift b ON a.employee_code = b.employee_code " & _
                "JOIN tm_shift c ON b.shift_number = c.shift_number " & _
              "WHERE a.company_code = '" & frm_list_manual_att.TDBCombo_company.Columns("company_code").Value & "' " & _
                "AND c.shift_code = '" & TDBCombo_shift.Columns("shift_code").Value & "' " & _
                "AND a.flag_active <> 0"
        rsemp.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        For aa = 0 To rsemp.RecordCount - 1
            If rscari.State Then rscari.Close
            SQL = "SELECT DISTINCT b.shift_number " & _
                  "FROM m_employee a JOIN td_shift b ON a.employee_code = b.employee_code " & _
                    "JOIN tm_shift c ON b.shift_number = c.shift_number " & _
                  "WHERE a.company_code = '" & frm_list_manual_att.TDBCombo_company.Columns("company_code").Value & "' " & _
                    "AND a.employee_code = '" & rsemp!employee_code & "' " & _
                    "AND c.shift_code = '" & TDBCombo_shift.Columns("shift_code").Value & "' " & _
                  "ORDER BY c.shift_date DESC LIMIT 1"
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
            If rscari.RecordCount > 0 Then
                vShiftNumber = IIf(IsNull(rscari!shift_number), 0, rscari!shift_number)
            Else
                vShiftNumber = 1
            End If
            rscari.Close
        
            If editTrans = False Then
                If rscari.State Then rscari.Close
                SQL = "SELECT enrollnumber FROM m_enroll_link WHERE employee_code = '" & rsemp!employee_code & "'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

                If rscari.RecordCount > 0 Then
                    vEnrollNumber = IIf(IsNull(rscari!enrollnumber), 0, rscari!enrollnumber)
                Else
                    vEnrollNumber = 0
                End If
                rscari.Close

                SQL = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglAwal, "yyyy-MM-dd HH:mm:ss") & "' AND employee_code = '" & rsemp!employee_code & "'"
                CnG.Execute SQL

                SQL = "INSERT INTO h_attendance (employee_code,att_date,enrollnumber,shift_code," & _
                        "shift_number,start_time,end_time,time_in,time_out,description,entry_date,entry_user," & _
                        "flag_present,flag_duty) " & _
                      "VALUES (" & _
                        "'" & rsemp!employee_code & "','" & Format(tglAwal, "yyyy-MM-dd") & "','" & vEnrollNumber & "'," & _
                        "'" & TDBCombo_shift.Text & "','" & vShiftNumber & "','" & start_time & "','" & end_time & "'," & _
                        "'" & time_in & "','" & time_out & "'," & _
                        "'" & txt_description.Text & "',now(),'" & LOGIN_NAME & "',1,0)"
            Else

                SQL = "UPDATE h_attendance set employee_code = '" & rsemp!employee_code & "'," & _
                        "enrollnumber = '" & vEnrollNumber & "'," & _
                        "shift_code = '" & TDBCombo_shift.Text & "'," & _
                        "shift_number = '" & vShiftNumber & "',start_time = '" & start_time & "',end_time = '" & end_time & "'," & _
                        "time_in = '" & time_in & "',flag_present = 1," & _
                        "time_out = '" & time_out & "'," & _
                        "description = '" & txt_description.Text & "',edit_date = now(),edit_user = '" & LOGIN_NAME & "'," & _
                        "flag_present = 1, flag_duty = 0 " & _
                        "WHERE employee_code = '" & rsemp!employee_code & "' " & _
                            "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd HH:mm:ss") & "'"
            End If
            CnG.Execute SQL
            rsemp.MoveNext
        Next
    Else
        If rscari.State Then rscari.Close
        SQL = "SELECT enrollnumber FROM m_enroll_link WHERE employee_code = '" & txt_employee_code.Text & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        If rscari.RecordCount > 0 Then
            vEnrollNumber = IIf(IsNull(rscari!enrollnumber), 0, rscari!enrollnumber)
        Else
            vEnrollNumber = 0
        End If
        rscari.Close
        
        If rscari.State Then rscari.Close
        SQL = "SELECT DISTINCT b.shift_number " & _
              "FROM m_employee a JOIN td_shift b ON a.employee_code = b.employee_code " & _
                "JOIN tm_shift c ON b.shift_number = c.shift_number " & _
              "WHERE a.company_code = '" & frm_list_manual_att.TDBCombo_company.Columns("company_code").Value & "' " & _
                "AND a.employee_code = '" & txt_employee_code.Text & "' " & _
                "AND c.shift_code = '" & TDBCombo_shift.Columns("shift_code").Value & "' " & _
              "ORDER BY c.shift_date DESC LIMIT 1"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        If rscari.RecordCount > 0 Then
            vShiftNumber = IIf(IsNull(rscari!shift_number), 0, rscari!shift_number)
        Else
            vShiftNumber = 1
        End If
        rscari.Close

        If editTrans = False Then
            SQL = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglAwal, "yyyy-MM-dd HH:mm:ss") & "' AND employee_code = '" & txt_employee_code.Text & "'"
            CnG.Execute SQL

            SQL = "INSERT INTO h_attendance (employee_code,att_date,enrollnumber,shift_code," & _
                      "shift_number,start_time,end_time,time_in,time_out,description,entry_date,entry_user," & _
                      "flag_present,flag_duty) " & _
                    "VALUES (" & _
                      "'" & txt_employee_code.Text & "','" & Format(tglAwal, "yyyy-MM-dd") & "','" & vEnrollNumber & "'," & _
                      "'" & TDBCombo_shift.Text & "','" & vShiftNumber & "','" & start_time & "','" & end_time & "'," & _
                      "'" & time_in & "','" & time_out & "'," & _
                      "'" & txt_description.Text & "',now(),'" & LOGIN_NAME & "',1,0)"
        Else
            SQL = "UPDATE h_attendance set employee_code = '" & txt_employee_code.Text & "'," & _
                    "enrollnumber = '" & vEnrollNumber & "'," & _
                    "shift_code = '" & TDBCombo_shift.Text & "'," & _
                    "shift_number = '" & vShiftNumber & "',start_time = '" & start_time & "',end_time = '" & end_time & "'," & _
                    "time_in = '" & time_in & "',flag_present = 1," & _
                    "time_out = '" & time_out & "'," & _
                    "description = '" & txt_description.Text & "',edit_date = now(),edit_user = '" & LOGIN_NAME & "'," & _
                    "flag_present = 1, flag_duty = 0 " & _
                    "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                        "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd HH:mm:ss") & "'"
        End If
        CnG.Execute SQL
    End If
End Sub

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
