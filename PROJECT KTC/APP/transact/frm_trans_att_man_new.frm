VERSION 5.00
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_att_man 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manual Attendance"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk_all_employee 
      Caption         =   "All Employee"
      Height          =   255
      Left            =   1290
      TabIndex        =   48
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtnmshift 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2220
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   46
      Top             =   930
      Width           =   2325
   End
   Begin prj_fej_jkt.LynxGrid LynxGrid2 
      Height          =   4575
      Left            =   1290
      TabIndex        =   38
      Top             =   1950
      Visible         =   0   'False
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   8070
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
   Begin VB.TextBox txtnmabsentstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   44
      Top             =   2400
      Width           =   2775
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frm_trans_att_man_new.frx":0000
      Left            =   8040
      List            =   "frm_trans_att_man_new.frx":0002
      TabIndex        =   3
      Text            =   "..."
      Top             =   2370
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtjmljam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8490
      TabIndex        =   40
      Top             =   2850
      Width           =   645
   End
   Begin TDBTime6Ctl.TDBTime ttin 
      Height          =   285
      Left            =   1290
      TabIndex        =   5
      Top             =   2850
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frm_trans_att_man_new.frx":0004
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frm_trans_att_man_new.frx":0070
      Spin            =   "frm_trans_att_man_new.frx":00C0
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
   Begin prj_fej_jkt.vbButton vbButton2 
      Height          =   285
      Left            =   2430
      TabIndex        =   37
      Top             =   1650
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
      MICON           =   "frm_trans_att_man_new.frx":00E8
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
      TabIndex        =   36
      Top             =   180
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   85917699
      CurrentDate     =   40823
   End
   Begin VB.TextBox txtnmDept 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frm_trans_att_man_new.frx":0104
      Left            =   3150
      List            =   "frm_trans_att_man_new.frx":0106
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2820
      Width           =   1695
   End
   Begin VB.TextBox txtkdshift 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6840
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   34
      Top             =   6390
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtnmshift1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   7710
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   33
      Top             =   6390
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   150
      TabIndex        =   31
      Top             =   4200
      Width           =   10695
   End
   Begin VB.TextBox txtkdtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1290
      TabIndex        =   30
      Top             =   1980
      Width           =   795
   End
   Begin VB.TextBox txtkddiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   7590
      TabIndex        =   29
      Top             =   1650
      Width           =   735
   End
   Begin VB.TextBox txtkdkar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1290
      TabIndex        =   0
      Top             =   1650
      Width           =   1125
   End
   Begin VB.TextBox txtnmkar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2790
      TabIndex        =   21
      Top             =   1650
      Width           =   4005
   End
   Begin VB.TextBox txtdivision 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8340
      TabIndex        =   20
      Top             =   1650
      Width           =   2535
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2100
      TabIndex        =   19
      Top             =   1980
      Width           =   3195
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   8
      Top             =   3210
      Width           =   9015
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   17
      Top             =   480
      Width           =   2025
   End
   Begin VB.TextBox txtentry 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   15
      Top             =   150
      Width           =   2025
   End
   Begin VB.ComboBox cmbdep 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4170
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   120
      Visible         =   0   'False
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
      TabIndex        =   11
      Top             =   570
      Width           =   3855
   End
   Begin TDBTime6Ctl.TDBTime ttout 
      Height          =   285
      Left            =   6270
      TabIndex        =   7
      Top             =   2850
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frm_trans_att_man_new.frx":0108
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frm_trans_att_man_new.frx":0174
      Spin            =   "frm_trans_att_man_new.frx":01C4
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
      MICON           =   "frm_trans_att_man_new.frx":01EC
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
      Left            =   1290
      OleObjectBlob   =   "frm_trans_att_man_new.frx":0208
      TabIndex        =   39
      Top             =   570
      Width           =   1785
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   6390
      TabIndex        =   2
      Top             =   2370
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   85917699
      CurrentDate     =   40823
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   315
      Left            =   8730
      TabIndex        =   4
      Top             =   2370
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   85917699
      CurrentDate     =   40823
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo1 
      Height          =   345
      Left            =   1290
      OleObjectBlob   =   "frm_trans_att_man_new.frx":216E
      TabIndex        =   1
      Top             =   2400
      Width           =   1185
   End
   Begin prj_fej_jkt.vbButton vbButton1 
      Height          =   585
      Left            =   675
      TabIndex        =   45
      Top             =   5175
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1032
      BTYPE           =   14
      TX              =   "Save"
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
      MICON           =   "frm_trans_att_man_new.frx":40BD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo2 
      Height          =   375
      Left            =   1290
      OleObjectBlob   =   "frm_trans_att_man_new.frx":40D9
      TabIndex        =   47
      Top             =   930
      Width           =   885
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   105
      Left            =   150
      TabIndex        =   49
      Top             =   4770
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label10 
      Caption         =   "Entry Date"
      Height          =   210
      Left            =   7830
      TabIndex        =   16
      Top             =   210
      Width           =   885
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   150
      TabIndex        =   50
      Top             =   4500
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label16 
      Caption         =   "From :"
      Height          =   195
      Left            =   5910
      TabIndex        =   43
      Top             =   2430
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Absen Status :"
      Height          =   195
      Left            =   210
      TabIndex        =   42
      Top             =   2460
      Width           =   1035
   End
   Begin VB.Label Label13 
      Caption         =   "Total Hour :"
      Height          =   210
      Left            =   7620
      TabIndex        =   41
      Top             =   2910
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Shift :"
      Height          =   195
      Left            =   840
      TabIndex        =   32
      Top             =   990
      Width           =   405
   End
   Begin VB.Label Label5 
      Caption         =   "Employee :"
      Height          =   210
      Left            =   450
      TabIndex        =   28
      Top             =   1710
      Width           =   795
   End
   Begin VB.Label Label6 
      Caption         =   "Division :"
      Height          =   210
      Left            =   6900
      TabIndex        =   27
      Top             =   1710
      Width           =   675
   End
   Begin VB.Label Label7 
      Caption         =   "Title :"
      Height          =   210
      Left            =   840
      TabIndex        =   26
      Top             =   2040
      Width           =   405
   End
   Begin VB.Label Label8 
      Caption         =   "Check IN :"
      Height          =   195
      Left            =   480
      TabIndex        =   25
      Top             =   2910
      Width           =   765
   End
   Begin VB.Label Label9 
      Caption         =   "Remark :"
      Height          =   210
      Left            =   630
      TabIndex        =   24
      Top             =   3270
      Width           =   645
   End
   Begin VB.Label Label12 
      Caption         =   "Break :"
      Height          =   195
      Left            =   2580
      TabIndex        =   23
      Top             =   2910
      Width           =   525
   End
   Begin VB.Label Label14 
      Caption         =   "Check OUT :"
      Height          =   195
      Left            =   5310
      TabIndex        =   22
      Top             =   2910
      Width           =   945
   End
   Begin VB.Label Label11 
      Caption         =   "User Name"
      Height          =   210
      Left            =   7830
      TabIndex        =   18
      Top             =   540
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Department :"
      Height          =   195
      Left            =   3210
      TabIndex        =   14
      Top             =   180
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Branch Office :"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   630
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Date :"
      Height          =   195
      Left            =   780
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frm_trans_att_man"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim strsql As String
Public editTrans As Boolean
Dim tgl_masuk As Date
Dim tgl_keluar As Date
Public v_tt_in, v_tt_out, v_absen_status As String
Dim tglawal As Date, tglAkhir As Date

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
Dim access As String

    access = IIf(LOGIN_LEVEL = 100, "", "AND (managerial_access = 0 OR managerial_access IS NULL)")
    
    If pilihan = 1 Then
        LynxGrid2.Clear
'        strsql = "select employee_code,employee_name,division_code,division_name," _
'                    & "title_code,title_name " _
'                & "from m_employee " _
'                & "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' AND " _
'                & "group_shift = '" & frm_list_manual_att.LynxGrid2.CellText(LynxGrid2.Row, 0) & "' " & access & " AND " _
'                & "(employee_code LIKE '%" & txtkdkar.Text & "%' " _
'                & "OR employee_name LIKE '%" & txtkdkar.Text & "%') ORDER BY employee_name"
        
        strsql = "select employee_code,employee_name,division_code,division_name," _
                    & "title_code,title_name " _
                & "from m_employee " _
                & "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' AND " _
                & "group_shift = '" & frm_list_manual_att.LynxGrid2.CellText(LynxGrid2.Row, 0) & "' " & access & " AND " _
                & "(employee_code LIKE '%" & txtkdkar.Text & "%' " _
                & "OR employee_name LIKE '%" & txtkdkar.Text & "%') " _
                & "AND employee_code not in (select employee_code from h_attendance WHERE date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "') " _
                & "ORDER BY employee_name"
                
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
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
                txtdivision.Text = IIf(IsNull(rs2!division_name) = True, "", rs2!division_name)
                txtkdtitle.Text = rs2!title_code
                txttitle.Text = rs2!title_name
                TDBCombo1.SetFocus
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
            TDBCombo1.SetFocus
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub chk_all_employee_Click()
    If chk_all_employee.Value = 0 Then
        txtkdkar.Enabled = True
        vbButton2.Enabled = True
    Else
        txtkdkar.Enabled = False
        vbButton2.Enabled = False
    End If
End Sub

Private Sub Combo1_Change()
    Call totalJam
End Sub

Private Sub Combo1_Click()
    Call totalJam
End Sub

Private Sub totalJam()
    Dim tglin As Date, tglout As Date
    Dim jambreak As Integer
    Dim tot_menit As Double
    Dim tgl_masuk As Date
    Dim tgl_keluar As Date
    Dim menit As Double
    Dim jam As Double
    
    Dim e As Long
    
    tglin = Date & " " & Format(ttin.Value, "hh:nn:ss")
    
    If Format(ttin.Value, "hh:nn:ss") > Format(ttout.Value, "hh:nn:s") Then
        tglout = Date + 1 & " " & ttout.Value
    Else
        tglout = Date & " " & ttout.Value
    End If
    
    Select Case Combo1.ListIndex
        Case 0
            jambreak = 1
        Case 1
            jambreak = 2
        Case 2
            jambreak = 0
    End Select
    
    tgl_masuk = Format(tglin, "yyyy-MM-dd hh:nn:ss")
    tgl_keluar = Format(tglout, "yyyy-MM-dd hh:nn:ss")
  
    menit = (DateDiff("n", tgl_masuk, tgl_keluar)) / 60
    jam = roundDown(menit)
'    jam = DateDiff("h", tgl_masuk, tgl_keluar)
    
    e = (menit - jam) * 60
    
    If e < 21 Then
        tot_menit = 0
    ElseIf e > 20 And e < 50 Then
        tot_menit = 0.5
    Else
        tot_menit = 1
    End If
    
txtjmljam.Text = jam - jambreak + tot_menit
End Sub

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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        vbButton1_Click
    End If
End Sub

Private Sub Form_Load()
Dim rsshift As New ADODB.Recordset
    createKar
    chk_all_employee.Value = 0
    
    Combo3.AddItem "..."
    Combo3.AddItem "To"
    Combo3.Text = "..."
    
'    DTPicker2.Value = Date
'    DTPicker3.Value = Date
    
'    If editTrans = False Then
'        isijamkerja
'    End If
    Call isiShift
    
    strsql = "select absent_code,absent_name from m_absent_status"
    rsshift.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo1.RowSource = rsshift
    
End Sub

Public Sub isijamkerja()
    With Combo1
        .AddItem "1 Jam Break"
        .AddItem "2 Jam Break"
        .AddItem "0 Jam Break"
    
        .ListIndex = 0
    End With
    If WeekdayName(Weekday(frm_list_manual_att.DTPicker1.Value)) = "Friday" Then
        Combo1.ListIndex = 1
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
            DTPicker2.Value = Format(DTPicker1.Value, "yyyy-MM-dd")
            DTPicker3.Value = Format(DTPicker1.Value, "yyyy-MM-dd")
        Else
            Label16.Visible = True
            DTPicker2.Visible = True
            Combo3.Visible = True
            ttin.Enabled = True
            ttout.Enabled = True
            Combo1.Enabled = True
            ttin.Text = "00:00"
            ttout.Text = "00:00"
            txtjmljam = "0"
            DTPicker2.Value = Format(DTPicker1.Value, "yyyy-MM-dd")
            DTPicker3.Value = Format(DTPicker1.Value, "yyyy-MM-dd")
        End If
    End If
End Sub

'Private Sub TDBCombo2_Click()
'    TDBCombo1.Text = "0"
'    txtnmabsentstatus.Text = "PRESENT"
'End Sub

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
'On Error GoTo err
On Error Resume Next

Dim rsemp As New ADODB.Recordset
Dim aa As Integer
Dim a, b, c, d As String
    If TDBCombo1.Text = "0" Then
        If ttin.Value = ttout.Value Then
            MsgBox "Invalid Check In Time...!!!", vbExclamation, "Error"
            Exit Sub
        End If
        
        If Format(ttout.Value, "hh:mm") = "00:00" Then
            MsgBox "Invalid Format Check Out Time....!!!", vbExclamation, "Error"
            Exit Sub
        End If
        
        If Format(ttout.Value, "hh:mm") = "" Or Format(ttout.Value, "hh:mm") = "" Then
            MsgBox "Invalid Format Check In or Check Out Time....!!!", vbExclamation, "Error"
            Exit Sub
        End If
    End If
    
    '+++++++++++++++++++CHECK SHIFT CODE++++++++++++++++++++++++++++++++++++
    If TDBCombo2.Text = "ALL" Then
        MsgBox "Please Select a Valid Shift Code!", vbExclamation, headerMSG
        TDBCombo2.SetFocus
        Exit Sub
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++APAKAH KODE KARYAWAN SUDAH BENAR++++++++++++++++++++
    If chk_all_employee.Value = 0 Then
        strsql = "SELECT 1 FROM m_employee WHERE employee_code = '" & txtkdkar.Text & "'"
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount = 0 Then
            MsgBox "Invalid Employee Code...!!", vbCritical
            rs2.Close
            Exit Sub
        End If
        rs2.Close
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++APAKAH KARYAWAN SUDAH PERNAH DI INPUT++++++++++++++++++++
    If editTrans = False Then
        If chk_all_employee.Value = 0 Then
            If TDBCombo1.Text = "0" Then 'Present
                strsql = "SELECT att_date,shift_code,shift_number FROM h_attendance WHERE employee_code = '" & txtkdkar.Text & "' " _
                    & "AND DATE(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
                rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
            ElseIf Combo3.ListIndex = 1 Then 'Lebih dr 1 hari
                strsql = "SELECT att_date,shift_code,shift_number FROM h_attendance WHERE employee_code = '" & txtkdkar.Text & "' " _
                    & "AND (DATE(att_date) BETWEEN '" & Format(DTPicker2.Value, "yyyy-MM-dd") & "' " _
                    & "AND '" & Format(DTPicker3.Value, "yyyy-MM-dd") & "')"
                rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
            Else
                strsql = "SELECT att_date,shift_code,shift_number FROM h_attendance WHERE employee_code = '" & txtkdkar.Text & "' " _
                    & "AND DATE(att_date) = '" & Format(DTPicker2.Value, "yyyy-MM-dd") & "'"
                rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
            End If
                    
            If rs2.RecordCount > 0 Then
                Dim v_shift As String
                v_shift = IIf(rs2!shift_code = rs2!shift_number, "PRESENT", IIf(rs2!shift_code = "A", "ALPHA", IIf(rs2!shift_code = "DT", "DUTY", _
                        IIf(rs2!shift_code = "OF", "OFF", IIf(rs2!shift_code = "PR", "CUTI", IIf(rs2!shift_code = "P", "Ijin / Sakit Tanpa Surat Dokter", _
                        "Sakit Ada Surat Dokter"))))))
                
                    MsgBox "This Employee Is Already Exist With Status " & v_shift, vbCritical
                    rs2.Close
                Exit Sub
            End If
            rs2.Close
        End If
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim v_flag_active As Integer
    tglawal = DTPicker2.Value
    If Combo3.Text = "To" Then
        tglAkhir = DTPicker3.Value
    Else
        tglAkhir = DTPicker2.Value
    End If
    
    strsql = "SELECT employee_code, flag_active FROM m_employee WHERE employee_code = '" & txtkdkar.Text & "'"
    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If Not rs2.EOF Then
            v_flag_active = rs2!flag_active
        End If
    rs2.Close
    
    If v_flag_active = 2 Then
        MsgBox "Employees in the MC condition, Update Attendance does not Allowed!"
        Exit Sub
    End If

    Select Case TDBCombo1.Text
        Case "0"
            CnG.BeginTrans
                                    
            While tglawal <= tglAkhir
                InsertPresent
                
                tglawal = tglawal + 1
            Wend
            
            ProgressBar1.Visible = False
            
            CnG.CommitTrans
        Case "A", "OF", "P", "S" ',"PR"
            Dim abstatus As Integer
            If TDBCombo1.Text = "A" Then
                abstatus = 2
            ElseIf TDBCombo1.Text = "OF" Then
                abstatus = 6
            ElseIf TDBCombo1.Text = "P" Then
                abstatus = 0
            ElseIf TDBCombo1.Text = "S" Then
                abstatus = 1
'            ElseIf TDBCombo1.Text = "PR" Then
'                abstatus = 3
            End If
            
            CnG.BeginTrans
            While tglawal <= tglAkhir
                
                If chk_all_employee.Value = 1 Then
                    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
                    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
                    CnG.Execute strsql
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
                    strsql = "SELECT employee_code, employee_name from m_employee " _
                            & "where company_code = '" & TDBCombo_company.Text & "' AND " _
                            & "group_shift = '" & cmbdep.Text & "'"
                    rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    ProgressBar1.Value = 0
                    ProgressBar1.Max = rsemp.RecordCount
                    ProgressBar1.Visible = True
                    lblDescription.Visible = True
                    
                    For aa = 0 To rsemp.RecordCount - 1
                        lblDescription = "(" & Format(tglawal, "dd-MM-yyyy") & ") - " & rsemp!employee_name & " - Inserting Data..."
                        ProgressBar1.Value = ProgressBar1.Value + 1
                        
                        DoEvents
                    
                        strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & rsemp!employee_code & "'"
                        CnG.Execute strsql
                        
                        strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_code,flag_present,absent_status,description,entry_date,userinput,shift_number) " _
                            & "VALUES " _
                            & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & rsemp!employee_code & "','" & TDBCombo1.Text & "',0,'" & abstatus & "','" & txtket.Text & "',now(),'" & LOGIN_CODE & "','" & TDBCombo2.Text & "')"
                        CnG.Execute strsql
                        rsemp.MoveNext
                    Next
                    rsemp.Close
                    
                    ProgressBar1.Visible = False
                    lblDescription.Visible = False
                Else
                    '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
                    If v_absen_status <> TDBCombo1.Text Then
                        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
                        CnG.Execute strsql
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                    strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & txtkdkar.Text & "'"
                    CnG.Execute strsql
                    
                    strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_code,flag_present,absent_status,description,entry_date,userinput,shift_number) " _
                        & "VALUES " _
                        & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & txtkdkar.Text & "','" & TDBCombo1.Text & "',0,'" & abstatus & "','" & txtket.Text & "',now(),'" & LOGIN_CODE & "','" & TDBCombo2.Text & "')"
                    CnG.Execute strsql
                End If
                
                tglawal = tglawal + 1
                
'                CnG.Execute "call spg_leave_periode2 ('" & Format(tglawal, "yyyy-MM-dd") & "')"
            Wend
            CnG.CommitTrans
        Case "DT"
            Dim waktu_in As String, waktu_out As String, waktu_break_out As String, waktu_break_in As String
            
            CnG.BeginTrans
            While tglawal <= tglAkhir
                
                    waktu_in = Format(tglawal, "yyyy-MM-dd") & " 07:00:00"
                    waktu_out = Format(tglawal, "yyyy-MM-dd") & " 17:00:00"
                    waktu_break_out = Format(tglawal, "yyyy-MM-dd") & " 12:00:00"
                    waktu_break_in = Format(tglawal, "yyyy-MM-dd") & " 12:00:00"
                
                If chk_all_employee.Value = 1 Then
                    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
                    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
                    CnG.Execute strsql
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    strsql = "SELECT employee_code, employee_name from m_employee " _
                            & "where company_code = '" & TDBCombo_company.Text & "' AND " _
                            & "group_shift= '" & cmbdep.Text & "'"
                    rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    ProgressBar1.Value = 0
                    ProgressBar1.Max = rsemp.RecordCount
                    ProgressBar1.Visible = True
                    lblDescription.Visible = True
                    
                    For aa = 0 To rsemp.RecordCount - 1
                        lblDescription = "(" & Format(tglawal, "dd-MM-yyyy") & ") - " & rsemp!employee_name & " - Inserting Data..."
                        ProgressBar1.Value = ProgressBar1.Value + 1
                        
                        DoEvents
                        
                        strsql = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & rsemp!employee_code & "'"
                        CnG.Execute strsql
                        
                        strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_code,time_in,time_out,time_out_break,time_in_break," _
                                & "flag_present,flag_duty,absent_status,description,entry_date,userinput,shift_number) " _
                            & "VALUES " _
                            & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & rsemp!employee_code & "','" & TDBCombo1.Text & "','" & waktu_in & "','" & waktu_out & "','" & waktu_break_out & "','" & waktu_break_in & "'," _
                                & "1,1,0,'" & txtket.Text & "',now(),'" & LOGIN_CODE & "','" & TDBCombo2.Text & "')"
                        CnG.Execute strsql
                        rsemp.MoveNext
                    Next
                    rsemp.Close
                    
                    ProgressBar1.Visible = False
                    lblDescription.Visible = False
                Else
                    '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
                    If v_absen_status <> TDBCombo1.Text Then
                        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
                        CnG.Execute strsql
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                    strsql = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & txtkdkar.Text & "'"
                    CnG.Execute strsql
                        
                    strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_code,time_in,time_out,time_out_break,time_in_break," _
                            & "flag_present,flag_duty,absent_status,description,entry_date,userinput,shift_number) " _
                        & "VALUES " _
                        & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & txtkdkar.Text & "','" & TDBCombo1.Text & "','" & waktu_in & "','" & waktu_out & "','" & waktu_break_out & "','" & waktu_break_in & "'," _
                            & "1,1,0,'" & txtket.Text & "',now(),'" & LOGIN_CODE & "','" & TDBCombo2.Text & "')"
                    CnG.Execute strsql
                End If
                
                tglawal = tglawal + 1
                
'                CnG.Execute "call spg_leave_periode2 ('" & Format(tglawal, "yyyy-MM-dd") & "')"
            Wend
            CnG.CommitTrans
            
        Case "PR"
            Dim rsleave As New ADODB.Recordset
            abstatus = 3
            
            CnG.BeginTrans
                CnG.Execute "call spg_leave_periode2 ('" & Format(tglawal, "yyyy-MM-dd") & "')"
                
                While tglawal <= tglAkhir
                strsql = "select actual_leave, max_leave from t_leave_periode " & _
                    "where employee_code = '" & txtkdkar.Text & "'"
                rsleave.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rsleave.RecordCount > 0 Then
                    If rsleave!actual_leave > rsleave!max_leave Then
                    MsgBox "Employee Leave is Over Limit..!" & Chr(13) & _
                        "Please Check Summary Leave!", vbExclamation, "Error!"
                        
                    Exit Sub
                    rsleave.Close
                    End If
                End If
                rsleave.Close
                
                If chk_all_employee.Value = 1 Then
                    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
                    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
                    CnG.Execute strsql
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                    strsql = "SELECT employee_code, employee_name from m_employee " _
                            & "where company_code = '" & TDBCombo_company.Text & "' AND " _
                            & "group_shift = '" & cmbdep.Text & "'"
                    rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    ProgressBar1.Value = 0
                    ProgressBar1.Max = rsemp.RecordCount
                    ProgressBar1.Visible = True
                    lblDescription.Visible = True
                    
                    For aa = 0 To rsemp.RecordCount - 1
                        lblDescription = "(" & Format(tglawal, "dd-MM-yyyy") & ") - " & rsemp!employee_name & " - Inserting Data..."
                        ProgressBar1.Value = ProgressBar1.Value + 1
                        
                        DoEvents
                        
                        strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & rsemp!employee_code & "'"
                        CnG.Execute strsql
                        
                        strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_code,flag_present,absent_status,description,entry_date,userinput,shift_number) " _
                            & "VALUES " _
                            & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & rsemp!employee_code & "','" & TDBCombo1.Text & "',0,'" & abstatus & "','" & txtket.Text & "',now(),'" & LOGIN_CODE & "','" & TDBCombo2.Text & "')"
                        CnG.Execute strsql
                        rsemp.MoveNext
                    Next
                    
                    ProgressBar1.Visible = False
                    lblDescription.Visible = False
                Else
                    '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
                    If v_absen_status <> TDBCombo1.Text Then
                        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
                        CnG.Execute strsql
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                    strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & txtkdkar.Text & "'"
                    CnG.Execute strsql
                    
                    strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_code,flag_present,absent_status,description,entry_date,userinput,shift_number) " _
                        & "VALUES " _
                        & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & txtkdkar.Text & "','" & TDBCombo1.Text & "',0,'" & abstatus & "','" & txtket.Text & "',now(),'" & LOGIN_CODE & "','" & TDBCombo2.Text & "')"
                    CnG.Execute strsql
                End If
                
                tglawal = tglawal + 1
                
'                CnG.Execute "call spg_leave_periode2 ('" & Format(tglawal, "yyyy-MM-dd") & "')"
            Wend
            CnG.CommitTrans
        Case Else
            MsgBox "Please Check Absent Status....!!!"
            Exit Sub
    End Select
    
    MsgBox "Save Data Succesfully!", vbInformation, "Sukses!"
    
    txtkdkar.Text = ""
    txtnmkar.Text = ""
    txtkddiv.Text = ""
    txtdivision.Text = ""
    txtkdtitle.Text = ""
    txttitle.Text = ""
    TDBCombo1.Text = ""
    txtnmabsentstatus.Text = ""
    ttin.Value = "00:00"
    Combo1.Text = "2 Jam Break"
    ttout.Value = "00:00"
    txtket.Text = ""
    If chk_all_employee.Value = 0 Then
        txtkdkar.SetFocus
    End If
    editTrans = False
    
    frm_list_manual_att.isiGridAbsen
    Exit Sub

err:
MsgBox "There Is Any Problem With Application!" & Chr(13) & _
    "Please Contact Us (PT. Solusi Sentral Data - (031) 5616465)", vbInformation, headerMSG
Exit Sub
End Sub

Private Sub vbButton2_Click()
    isiGridKar (1)
End Sub

Private Sub vbButton5_Click()
    Unload Me
End Sub

Private Sub InsertPresent()
Dim start_time As String, end_time As String, max_break_out As String, min_break_in As String
Dim time_in As String, time_out_break As String, time_in_break As String, time_out As String
Dim rsemp As New ADODB.Recordset
   
    '+++++++++++++++++++++MENCARI TANGGAL BATAS JAM MASUK,ISTIRAHAT,KELUAR++++++++
    strsql = "Select CAST(concat('" & Format(tglawal, "yyyy-MM-dd") & "',' ', time(start_time)) as datetime) start_time," _
            & "CAST(concat('" & Format(tglawal, "yyyy-MM-dd") & "',' ', time(end_time)) as datetime) end_time," _
            & "CAST(concat('" & Format(tglawal, "yyyy-MM-dd") & "',' ', time(max_break_out)) as datetime) max_break_out," _
            & "CAST(concat('" & Format(tglawal, "yyyy-MM-dd") & "',' ', time(min_break_in)) as datetime) min_break_in," _
            & "curdate() tglserver " _
            & "from m_shift where shift_code = '" & TDBCombo2.Text & "'"
'        & "from m_shift where shift_code = '" & txtkdshift.Text & "'"
    rs2.Open strsql, CnG, adOpenDynamic, adLockReadOnly
    
    start_time = Format(rs2!start_time, "yyyy-MM-dd hh:mm:ss")
    If TDBCombo2.Text = "NS" Or TDBCombo2.Text = "NSS" Then
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
    
    time_in = Format(tglawal, "yyyy-MM-dd") & " " & Format(ttin.Value, "hh:mm") & ":00"
    
    time_out_break = max_break_out
    time_in_break = min_break_in
    
    If Format(ttout.Value, "hh:mm:ss") < Format(ttin.Value, "hh:mm:ss") Then
        time_out = Format(tglawal + 1, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    Else
        time_out = Format(tglawal, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    End If
    
    rs2.Close
    
    If chk_all_employee.Value = 1 Then
        strsql = "SELECT employee_code, employee_name from m_employee " _
                & "where company_code = '" & TDBCombo_company.Text & "' AND " _
                & "group_shift = '" & frm_list_manual_att.LynxGrid2.CellText(LynxGrid2.Row, 0) & "'"
        rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        ProgressBar1.Value = 0
        ProgressBar1.Max = rsemp.RecordCount
        ProgressBar1.Visible = True
        lblDescription.Visible = True
        
        For aa = 0 To rsemp.RecordCount - 1
            If editTrans = False Then
                lblDescription = "(" & Format(tglawal, "dd-MM-yyyy") & ") - " & rsemp!employee_name & " - Inserting Data..."
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                DoEvents
                
                '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
                strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
                CnG.Execute strsql
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' AND employee_code = '" & rsemp!employee_code & "'"
                CnG.Execute strsql
                        
                strsql = "INSERT INTO h_attendance (employee_code,att_date,shift_code,shift_number,start_time,end_time," _
                        & "time_in," _
                        & "time_out_break,time_out_break_diff," _
                        & "time_in_break,time_in_break_diff," _
                        & "time_out,description,entry_date,userinput,absent_status,flag_present,flag_duty) VALUES " _
                        & "('" & rsemp!employee_code & "','" & Format(tglawal, "yyyy-MM-dd") & "','" & TDBCombo2.Text & "','" & TDBCombo2.Text & "','" & start_time & "','" & end_time & "'," _
                        & "'" & time_in & "'," _
                        & "'" & time_out_break & "',CASE WHEN TIMEDIFF('" & time_out_break & "','" & max_break_out & "') < 0 THEN TIMEDIFF('" & max_break_out & "','" & time_out_break & "') ELSE TIMEDIFF('" & time_out_break & "','" & max_break_out & "') END," _
                        & "'" & time_in_break & "',CASE WHEN TIMEDIFF('" & time_in_break & "','" & min_break_in & "') < 0 THEN TIMEDIFF('" & min_break_in & "','" & time_in_break & "') ELSE TIMEDIFF('" & time_in_break & "','" & min_break_in & "') END," _
                        & "'" & time_out & "'," _
                        & "'" & txtket.Text & "',now(),'" & LOGIN_CODE & "',0,1,0)"
            Else
                lblDescription = "(" & Format(tglawal, "dd-MM-yyyy") & ") - " & rsemp!employee_name & " - Inserting Data..."
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                DoEvents
                
                '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
                If v_absen_status <> TDBCombo1.Text Then
                    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
                    CnG.Execute strsql
                ElseIf v_tt_in <> ttin.Value Or v_tt_out <> ttout.Value Then
                    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
                    CnG.Execute strsql
                End If
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                strsql = "UPDATE h_attendance set employee_code = '" & rsemp!employee_code & "'," _
                        & "att_date = '" & Format(tglawal, "yyyy-MM-dd") & "',shift_code = '" & TDBCombo2.Text & "'," _
                        & "shift_number = '" & TDBCombo2.Text & "',start_time = '" & start_time & "',end_time = '" & end_time & "'," _
                        & "time_in = '" & time_in & "',flag_present = 1," _
                        & "time_out_break = '" & time_out_break & "'," _
                        & "time_out_break_diff = CASE WHEN TIMEDIFF('" & time_out_break & "','" & max_break_out & "') < 0 THEN TIMEDIFF('" & max_break_out & "','" & time_out_break & "') ELSE TIMEDIFF('" & time_out_break & "','" & max_break_out & "') END," _
                        & "time_in_break = '" & time_in_break & "'," _
                        & "time_in_break_diff = CASE WHEN TIMEDIFF('" & time_in_break & "','" & min_break_in & "') < 0 THEN TIMEDIFF('" & min_break_in & "','" & time_in_break & "') ELSE TIMEDIFF('" & time_in_break & "','" & min_break_in & "') END," _
                        & "time_out = '" & time_out & "'," _
                        & "description = '" & txtket.Text & "',edit_date = now(),useredit = '" & LOGIN_CODE & "',absent_status = 0, flag_present = 1, flag_duty = 0 " _
                        & "WHERE employee_code = '" & txtkdkar.Text & "' " _
                        & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
        '                & "AND shift_code = '" & txtkdshift.Text & "'"
            End If
            CnG.Execute strsql
            rsemp.MoveNext
        Next
        
        ProgressBar1.Visible = False
        lblDescription.Visible = False
    Else
        
        If editTrans = False Then
            '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
            strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
            CnG.Execute strsql
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            strsql = "INSERT INTO h_attendance (employee_code,att_date,shift_code,shift_number,start_time,end_time," _
                    & "time_in," _
                    & "time_out_break,time_out_break_diff," _
                    & "time_in_break,time_in_break_diff," _
                    & "time_out,description,entry_date,userinput,absent_status,flag_present,flag_duty) VALUES " _
                    & "('" & txtkdkar.Text & "','" & Format(tglawal, "yyyy-MM-dd") & "','" & TDBCombo2.Text & "','" & TDBCombo2.Text & "','" & start_time & "','" & end_time & "'," _
                    & "'" & time_in & "'," _
                    & "'" & time_out_break & "',CASE WHEN TIMEDIFF('" & time_out_break & "','" & max_break_out & "') < 0 THEN TIMEDIFF('" & max_break_out & "','" & time_out_break & "') ELSE TIMEDIFF('" & time_out_break & "','" & max_break_out & "') END," _
                    & "'" & time_in_break & "',CASE WHEN TIMEDIFF('" & time_in_break & "','" & min_break_in & "') < 0 THEN TIMEDIFF('" & min_break_in & "','" & time_in_break & "') ELSE TIMEDIFF('" & time_in_break & "','" & min_break_in & "') END," _
                    & "'" & time_out & "'," _
                    & "'" & txtket.Text & "',now(),'" & LOGIN_CODE & "',0,1,0)"
        Else
            '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
            If v_absen_status <> TDBCombo1.Text Then
                strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
                CnG.Execute strsql
            ElseIf v_tt_in <> ttin.Value Or v_tt_out <> ttout.Value Then
                strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
                CnG.Execute strsql
            End If
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                
            strsql = "UPDATE h_attendance set employee_code = '" & txtkdkar.Text & "'," _
                    & "att_date = '" & Format(tglawal, "yyyy-MM-dd") & "',shift_code = '" & TDBCombo2.Text & "'," _
                    & "shift_number = '" & TDBCombo2.Text & "',start_time = '" & start_time & "',end_time = '" & end_time & "'," _
                    & "time_in = '" & time_in & "',flag_present = 1," _
                    & "time_out_break = '" & time_out_break & "'," _
                    & "time_out_break_diff = CASE WHEN TIMEDIFF('" & time_out_break & "','" & max_break_out & "') < 0 THEN TIMEDIFF('" & max_break_out & "','" & time_out_break & "') ELSE TIMEDIFF('" & time_out_break & "','" & max_break_out & "') END," _
                    & "time_in_break = '" & time_in_break & "'," _
                    & "time_in_break_diff = CASE WHEN TIMEDIFF('" & time_in_break & "','" & min_break_in & "') < 0 THEN TIMEDIFF('" & min_break_in & "','" & time_in_break & "') ELSE TIMEDIFF('" & time_in_break & "','" & min_break_in & "') END," _
                    & "time_out = '" & time_out & "'," _
                    & "description = '" & txtket.Text & "',edit_date = now(),useredit = '" & LOGIN_CODE & "',absent_status = 0, flag_present = 1, flag_duty = 0 " _
                    & "WHERE employee_code = '" & txtkdkar.Text & "' " _
                    & "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
    '                & "AND shift_code = '" & txtkdshift.Text & "'"
        End If
        CnG.Execute strsql
    End If
    
'    CnG.Execute "call spg_leave_periode2 ('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "')"

End Sub

Private Sub isiShift()
Dim rsshift As New ADODB.Recordset
    strsql = "select shift_code,shift_name " _
            & "from m_shift"
    rsshift.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo2.RowSource = rsshift
End Sub

Private Sub TDBCombo2_ItemChange()
    If TDBCombo2.ApproxCount > 0 Then
        TDBCombo2.Text = TDBCombo2.Columns("shift_code").Value
        txtnmshift.Text = TDBCombo2.Columns("shift_name").Value
    End If
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
