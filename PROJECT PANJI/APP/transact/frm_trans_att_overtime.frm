VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_overtime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manual Attendance"
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
   Begin prj_hardys.LynxGrid LynxGrid2 
      Height          =   4605
      Left            =   1290
      TabIndex        =   29
      Top             =   1740
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8123
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
   Begin prj_hardys.vbButton vbButton2 
      Height          =   285
      Left            =   2430
      TabIndex        =   28
      Top             =   1440
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
      MICON           =   "frm_trans_att_overtime.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtnmDept 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   930
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   150
      TabIndex        =   23
      Top             =   4530
      Width           =   10695
   End
   Begin VB.TextBox txtkdtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1290
      TabIndex        =   22
      Top             =   1770
      Width           =   795
   End
   Begin VB.TextBox txtkddiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   7590
      TabIndex        =   21
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtkdkar 
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
      TabIndex        =   15
      Top             =   1440
      Width           =   4005
   End
   Begin VB.TextBox txtdivision 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8340
      TabIndex        =   14
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2100
      TabIndex        =   13
      Top             =   1770
      Width           =   3195
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2490
      Width           =   9015
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   11
      Top             =   480
      Width           =   2025
   End
   Begin VB.TextBox txtentry 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   9
      Top             =   150
      Width           =   2025
   End
   Begin VB.ComboBox cmbdep 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1290
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   930
      Width           =   1095
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   570
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1290
      TabIndex        =   2
      Top             =   180
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   96927745
      CurrentDate     =   40794
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1290
      OleObjectBlob   =   "frm_trans_att_overtime.frx":001C
      TabIndex        =   5
      Top             =   570
      Width           =   1785
   End
   Begin prj_hardys.vbButton vbButton1 
      Height          =   615
      Left            =   750
      TabIndex        =   30
      Top             =   5160
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1085
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
      MICON           =   "frm_trans_att_overtime.frx":1FDA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_hardys.vbButton vbButton5 
      Height          =   615
      Left            =   2550
      TabIndex        =   31
      Top             =   5160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
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
      MICON           =   "frm_trans_att_overtime.frx":1FF6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_hardys.EnterNum EnterNum1 
      Height          =   315
      Left            =   1290
      TabIndex        =   32
      Top             =   2130
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
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
   Begin VB.Label Label12 
      Caption         =   "Jam"
      Height          =   195
      Left            =   2730
      TabIndex        =   33
      Top             =   2190
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Work Status :"
      Height          =   195
      Left            =   5640
      TabIndex        =   24
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Employee :"
      Height          =   210
      Left            =   450
      TabIndex        =   20
      Top             =   1500
      Width           =   795
   End
   Begin VB.Label Label6 
      Caption         =   "Division :"
      Height          =   210
      Left            =   6900
      TabIndex        =   19
      Top             =   1500
      Width           =   675
   End
   Begin VB.Label Label7 
      Caption         =   "Title :"
      Height          =   210
      Left            =   840
      TabIndex        =   18
      Top             =   1830
      Width           =   405
   End
   Begin VB.Label Label8 
      Caption         =   "Lembur :"
      Height          =   195
      Left            =   630
      TabIndex        =   17
      Top             =   2190
      Width           =   645
   End
   Begin VB.Label Label9 
      Caption         =   "Remark :"
      Height          =   210
      Left            =   630
      TabIndex        =   16
      Top             =   2550
      Width           =   645
   End
   Begin VB.Label Label11 
      Caption         =   "User Name"
      Height          =   210
      Left            =   7830
      TabIndex        =   12
      Top             =   540
      Width           =   885
   End
   Begin VB.Label Label10 
      Caption         =   "Entry Date"
      Height          =   210
      Left            =   7830
      TabIndex        =   10
      Top             =   210
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Department :"
      Height          =   195
      Left            =   330
      TabIndex        =   8
      Top             =   990
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Branch Office :"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   630
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Date :"
      Height          =   195
      Left            =   780
      TabIndex        =   3
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
Dim strSQL As String
Public editTrans As Boolean

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
                & "WHERE a.company_code = '" & TDBCombo_company.Text & "' AND " _
                & "a.department_code = '" & cmbdep.Text & "' AND " _
                & "(a.employee_code LIKE '%" & txtkdkar.Text & "%' " _
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
    If KeyAscii = 13 Then
        SendKeys (vbTab)
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        vbButton1_Click
    End If
End Sub

Private Sub Form_Load()
    createKar
    With Combo1
        .AddItem "1 Jam Break"
        .AddItem "2 Jam Break"
        .AddItem "0 Jam Break"
    
        .Text = "1 Jam Break"
    End With
    If WeekdayName(Weekday(DTPicker1.Value)) = "Friday" Then
        Combo1.Text = "2 Jam Break"
    End If
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
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub LynxGrid2_LostFocus()
    LynxGrid2.Visible = False
End Sub

Private Sub ttin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys (vbTab)
    End If
End Sub

Private Sub ttout_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys (vbTab)
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
Dim start_time As String, end_time As String, max_break_out As String, min_break_in As String
Dim time_in As String, time_out_break As String, time_in_break As String, time_out As String
Dim check_in_diff As Double
Dim flagCheckIn_Early As Integer, flagCheckIn_Late As Integer

    If ttin.Value = ttout.Value Then
        MsgBox "Invalid Check In Time...!!!", vbExclamation, "Error"
        Exit Sub
    End If
    '+++++++++++++++++++APAKAH KODE KARYAWAN SUDAH BENAR++++++++++++++++++++
    strSQL = "SELECT 1 FROM m_employee WHERE employee_code = '" & txtkdkar.Text & "'"
    rs2.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rs2.RecordCount = 0 Then
        MsgBox "Invalid Employee Code...!!", vbCritical
        rs2.Close
        Exit Sub
    End If
    rs2.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++APAKAH KARYAWAN SUDAH PERNAH DI INPUT++++++++++++++++++++
    If editTrans = False Then
        strSQL = "SELECT 1 FROM h_attendance WHERE employee_code = '" & txtkdkar.Text & "' " _
            & "AND DATE(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
        rs2.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            MsgBox "This Employee Is Already Exist...!!", vbCritical
            rs2.Close
            Exit Sub
        End If
        rs2.Close
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++++MENCARI TANGGAL BATAS JAM MASUK,ISTIRAHAT,KELUAR++++++++
    strSQL = "Select CAST(concat('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',' ', time(start_time)) as datetime) start_time," _
            & "CAST(concat('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',' ', time(end_time)) as datetime) end_time," _
            & "CAST(concat('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',' ', time(max_break_out)) as datetime) max_break_out," _
            & "CAST(concat('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',' ', time(min_break_in)) as datetime) min_break_in," _
            & "curdate() tglserver " _
        & "from m_shift where shift_code = '" & txtkdshift.Text & "'"
    rs2.Open strSQL, CnG, adOpenDynamic, adLockReadOnly
    
    start_time = Format(rs2!start_time, "yyyy-MM-dd hh:mm:ss")
    If txtkdshift.Text = "NS" Or txtkdshift.Text = "NSS" Then
        end_time = Format(DateAdd("d", 1, rs2!end_time), "yyyy-MM-dd hh:mm:ss")
    Else
        end_time = Format(rs2!end_time, "yyyy-MM-dd hh:mm:ss")
    End If
    
    If rs2!max_break_out < rs2!start_time Then
        max_break_out = Format(rs2!max_break_out + 1, "yyyy-MM-dd hh:mm:ss")
    Else
        max_break_out = Format(rs2!max_break_out, "yyyy-MM-dd hh:mm:ss")
    End If
    
    If rs2!min_break_in < rs2!start_time Then
        min_break_in = Format(rs2!min_break_in + 1, "yyyy-MM-dd hh:mm:ss")
    Else
        min_break_in = Format(rs2!min_break_in, "yyyy-MM-dd hh:mm:ss")
    End If
        
    Select Case Left(Combo1.Text, 1)
    Case "2"
        min_break_in = Format(DateAdd("h", 2, max_break_out), "yyyy-MM-dd hh:mm:ss")
    Case "1"
        min_break_in = Format(DateAdd("h", 1, max_break_out), "yyyy-MM-dd hh:mm:ss")
    Case Else
        min_break_in = max_break_out
    End Select
    
    time_in = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(ttin.Value, "hh:mm") & ":00"
    
'    If Format(ttbreakout.Value, "hh:mm") < Format(start_time, "hh:mm") Then
'        time_out_break = Format(rs2!tglserver + 1, "yyyy-MM-dd") & " " & Format(ttbreakout.Value, "hh:mm") & ":00"
'    Else
'        time_out_break = Format(rs2!tglserver, "yyyy-MM-dd") & " " & Format(ttbreakout.Value, "hh:mm") & ":00"
'    End If
'    If Format(ttbreakin.Value, "hh:mm") < Format(start_time, "hh:mm") Then
'        time_in_break = Format(rs2!tglserver + 1, "yyyy-MM-dd") & " " & Format(ttbreakin.Value, "hh:mm") & ":00"
'    Else
'        time_in_break = Format(rs2!tglserver, "yyyy-MM-dd") & " " & Format(ttbreakin.Value, "hh:mm") & ":00"
'    End If
    
    time_out_break = max_break_out
    time_in_break = min_break_in

    If Format(ttout.Value, "hh:mm") < Format(start_time, "hh:mm") Then
        time_out = Format(DTPicker1.Value + 1, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    Else
        time_out = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    End If
    
    check_in_diff = DateDiff("h", start_time, time_in)
    
    If check_in_diff > 0 Then
            flagCheckIn_Late = 1
    ElseIf check_in_diff < 0 Then
            flagCheckIn_Early = 1
    End If
    
    rs2.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If editTrans = False Then
        strSQL = "INSERT INTO h_attendance (employee_code,att_date,shift_code,shift_number,start_time,end_time," _
                & "flag_io,flag_present,time_in,time_in_diff," _
                & "time_out_break,time_out_break_diff," _
                & "time_in_break,time_in_break_diff," _
                & "time_out,time_out_diff," _
                & "flag_late, flag_early, description,entry_date,userinput) VALUES " _
                & "('" & txtkdkar.Text & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & txtkdshift.Text & "',1,'" & start_time & "','" & end_time & "'," _
                & "0,1,'" & time_in & "',TIMEDIFF('" & time_in & "','" & start_time & "')," _
                & "'" & time_out_break & "',TIMEDIFF('" & time_out_break & "','" & max_break_out & "')," _
                & "'" & time_in_break & "',TIMEDIFF('" & time_in_break & "','" & min_break_in & "')," _
                & "'" & time_out & "',TIMEDIFF('" & time_out & "','" & end_time & "')," _
                & "'" & flagCheckIn_Late & "','" & flagCheckIn_Early & "'," _
                & "'" & txtket.Text & "',now(),'" & LOGIN_CODE & "')"
    Else
        strSQL = "UPDATE h_attendance set employee_code = '" & txtkdkar.Text & "'," _
                & "att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',shift_code = '" & txtkdshift.Text & "'," _
                & "shift_number = 1,start_time = '" & start_time & "',end_time = '" & end_time & "'," _
                & "flag_io = 0, " _
                & "flag_present = 1, " _
                & "time_in = '" & time_in & "'," _
                & "time_in_diff = TIMEDIFF('" & time_in & "','" & start_time & "')," _
                & "time_out_break = '" & time_out_break & "'," _
                & "time_out_break_diff = TIMEDIFF('" & time_out_break & "','" & max_break_out & "')," _
                & "time_in_break = '" & time_in_break & "'," _
                & "time_in_break_diff = TIMEDIFF('" & time_in_break & "','" & min_break_in & "')," _
                & "time_out = '" & time_out & "'," _
                & "time_out_diff = TIMEDIFF('" & time_out & "','" & end_time & "')," _
                & "flag_late = '" & flagCheckIn_Late & "'," _
                & "flag_early = '" & flagCheckIn_Early & "'," _
                & "description = '" & txtket.Text & "',edit_date = now(),useredit = '" & LOGIN_CODE & "' " _
                & "WHERE employee_code = '" & txtkdkar.Text & "' " _
                & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " _
                & "AND shift_code = '" & txtkdshift.Text & "'"
    End If
    CnG.Execute strSQL
    
    txtkdkar.Text = ""
    txtnmkar.Text = ""
    txtkddiv.Text = ""
    txtdivision.Text = ""
    txtkdtitle.Text = ""
    txttitle.Text = ""
    ttin.Value = "00:00"
    Combo1.Text = "2 Jam Break"
    ttout.Value = "00:00"
    txtket.Text = ""
    txtkdkar.SetFocus
    
    frm_list_manual_att.isiGridAbsen
End Sub

Private Sub vbButton2_Click()
    isiGridKar (1)
    LynxGrid2.Visible = True
End Sub

Private Sub vbButton5_Click()
    Unload Me
End Sub
