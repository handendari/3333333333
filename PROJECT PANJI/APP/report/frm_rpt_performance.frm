VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_rpt_performance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LAPORAN - PERFORMA KARYAWAN"
   ClientHeight    =   7545
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   10560
   Icon            =   "frm_rpt_performance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin prj_panji.LynxGrid LynxGrid2 
      Height          =   2925
      Left            =   4710
      TabIndex        =   26
      Top             =   3930
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5159
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
      GridLines       =   2
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      AllowColumnResizing=   -1  'True
      ColumnSort      =   -1  'True
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3330
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1110
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   1620
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "TAHUNAN"
      TabPicture(0)   =   "frm_rpt_performance.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   840
         TabIndex        =   5
         Top             =   810
         Width           =   8355
         Begin VB.ComboBox cbo_monthly_company 
            Height          =   315
            ItemData        =   "frm_rpt_performance.frx":05A6
            Left            =   3600
            List            =   "frm_rpt_performance.frx":05B0
            TabIndex        =   11
            Text            =   "..."
            Top             =   1590
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox cbo_yearly_employee 
            Height          =   315
            ItemData        =   "frm_rpt_performance.frx":05C3
            Left            =   1800
            List            =   "frm_rpt_performance.frx":05CD
            TabIndex        =   10
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin VB.ComboBox cbo_yearly_division 
            Height          =   315
            ItemData        =   "frm_rpt_performance.frx":05DE
            Left            =   1800
            List            =   "frm_rpt_performance.frx":05E8
            TabIndex        =   9
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.Frame fra_yearly_department 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   3600
            TabIndex        =   6
            Top             =   750
            Width           =   4695
            Begin VB.TextBox txt_yearly_division_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               MultiLine       =   -1  'True
               TabIndex        =   7
               Top             =   90
               Width           =   2415
            End
            Begin TrueOleDBList60.TDBCombo TDBCombo_yearly_division 
               Height          =   375
               Left            =   30
               OleObjectBlob   =   "frm_rpt_performance.frx":05FB
               TabIndex        =   8
               Top             =   90
               Width           =   1815
            End
            Begin MSAdodcLib.Adodc Adodc_yearly_division 
               Height          =   375
               Left            =   690
               Top             =   90
               Visible         =   0   'False
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
         End
         Begin MSComCtl2.DTPicker DTPicker_yearly 
            Height          =   300
            Left            =   1800
            TabIndex        =   12
            Top             =   1590
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   95223811
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Frame fra_yearly_employee 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   3630
            TabIndex        =   13
            Top             =   1050
            Width           =   4575
            Begin VB.TextBox txt_yearly_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               DragMode        =   1  'Automatic
               Height          =   285
               Left            =   1920
               TabIndex        =   24
               Top             =   150
               Width           =   2385
            End
            Begin VB.TextBox txt_yearly_nik 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   0
               TabIndex        =   23
               Top             =   150
               Width           =   1515
            End
            Begin VB.TextBox txt_yearly_employee_code 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4170
               TabIndex        =   22
               Top             =   150
               Visible         =   0   'False
               Width           =   255
            End
            Begin prj_panji.vbButton cmd_monthly_browse_employee 
               Height          =   285
               Left            =   1560
               TabIndex        =   25
               Top             =   150
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
               MICON           =   "frm_rpt_performance.frx":2569
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KARYAWAN"
            Height          =   195
            Left            =   750
            TabIndex        =   16
            Top             =   1230
            Width           =   930
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TAHUN"
            Height          =   195
            Left            =   1125
            TabIndex        =   15
            Top             =   1620
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DIVISI"
            Height          =   195
            Left            =   1230
            TabIndex        =   14
            Top             =   870
            Width           =   465
         End
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1530
      OleObjectBlob   =   "frm_rpt_performance.frx":2585
      TabIndex        =   3
      Top             =   1110
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1350
      Top             =   1140
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9000
      Top             =   1140
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   6060
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   840
         Top             =   360
      End
      Begin prj_panji.vbButton cmdExit 
         Height          =   705
         Left            =   8730
         TabIndex        =   19
         Top             =   300
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Keluar"
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
         MICON           =   "frm_rpt_performance.frx":44EB
         PICN            =   "frm_rpt_performance.frx":4507
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdPrint 
         Height          =   705
         Left            =   4860
         TabIndex        =   20
         Top             =   270
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Detail"
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
         MICON           =   "frm_rpt_performance.frx":5599
         PICN            =   "frm_rpt_performance.frx":55B5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdSummary 
         Height          =   705
         Left            =   3810
         TabIndex        =   21
         Top             =   270
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Rekap"
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
         MICON           =   "frm_rpt_performance.frx":6647
         PICN            =   "frm_rpt_performance.frx":6663
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LAPORAN PERFORMA KARYAWAN"
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
      Left            =   240
      TabIndex        =   18
      Top             =   150
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_rpt_performance.frx":76F5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12690
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "LAPORAN PPH PASAL 21"
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
      Left            =   240
      TabIndex        =   17
      Top             =   180
      Width           =   3285
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1170
      Width           =   1110
   End
End
Attribute VB_Name = "frm_rpt_performance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim int_libur() As Integer

Private Sub cbo_yearly_division_Click()
    TDBCombo_yearly_division.Text = ""
    txt_yearly_division_name.Text = ""
        
    If cbo_yearly_division.ListIndex = 0 Then
        fra_yearly_department.Visible = False
    Else
        fra_yearly_department.Visible = True
    End If
End Sub

Private Sub cbo_yearly_employee_Click()
    If cbo_yearly_employee.ListIndex = 0 Then
        fra_yearly_employee.Visible = False
    Else
        fra_yearly_employee.Visible = True
    End If
    
    txt_yearly_employee_code = "": txt_yearly_employee_name = "": txt_yearly_nik = ""
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSummary_Click()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2 As String
    
    int_process = vbNo
        
    str_file = "\report\rpt_summary_perf.rpt"
    
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
    int_flag_employee = 0
    
'    d1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
'    d2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
    
    If cbo_yearly_employee.Text = "All" Then
        str_employee_code = ""
    Else
        str_employee_code = txt_yearly_employee_code.Text
    End If
    
'    d1 = Format(DTPicker_monthly.Value, "yyyy-MM") & "-01"
'    d2 = Format(DTPicker_monthly.Value, "yyyy-MM") & getEndDay(Format(DTPicker_monthly.Value, "MM"), Format(DTPicker_monthly.Value, "yyyy"))
    
    str_sql = "SELECT a.company_code, e.company_name, a.division_code, f.division_name, " & _
                "a.employee_code, a.nik, a.employee_name, b.perf_avg, c.perf_grade, '" & IIf(LOGIN_LEVEL = 100, LOGIN_FULLNAME, EMPLOYEE_NAME) & "' " & _
              "FROM m_employee a join t_employee_perf_result b on a.employee_code = b.employee_code " & _
                "JOIN m_performance c on b.perf_number = c.perf_number " & _
                "JOIN m_company e on a.company_code = e.company_code " & _
                "JOIN m_division f on a.company_code = f.company_code AND a.division_code = f.division_code " & _
              "WHERE " & IIf(cbo_yearly_division.ListIndex = 0, "a.company_code = '" & TDBCombo_company.Text & "' ", "a.company_code = '" & TDBCombo_company.Text & "' AND a.division_code = '" & TDBCombo_yearly_division.Text & "' ") & " " & _
                "" & IIf(cbo_yearly_employee.ListIndex = 0, "", "AND a.employee_code = '" & txt_yearly_employee_code.Text & " '") & " " & _
                "AND b.perf_year = '" & Format(DTPicker_yearly.Value, "yyyy") & "' " & _
              "ORDER BY a.employee_name"
    str_param_periode = Format(DTPicker_yearly, "yyyy-MM")
    
    Call a.Show
    a.Caption = "REKAPITULASI PERFORMA KARYAWAN"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
    End Sub

Private Function check_valid_yearly() As Boolean
    check_valid_yearly = True
    
    'validate employee
    If cbo_yearly_employee.ListIndex = 1 Then
        If Trim(txt_yearly_employee_code) = "" Then
            MsgBox "Karyawan Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            cmd_monthly_browse_employee.SetFocus
            check_valid_yearly = False
            Exit Function
        End If
    End If
End Function

Private Sub rpt_yearly()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2 As String

    str_file = "\report\rpt_detail_perf.rpt"
    
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
    int_flag_employee = 0
    
    If cbo_yearly_employee.Text = "All" Then
        str_employee_code = ""
    Else
        str_employee_code = txt_yearly_employee_code.Text
    End If
    
    d1 = Format(DTPicker_yearly.Value, "yyyy-01-01")
    d2 = Format(DTPicker_yearly.Value, "yyyy-12-27")
    
'    str_sql = "CALL spr_pph21('" & d1 & "','" & d2 & "'," _
'            & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_yearly_division.Text & "'," _
'            & int_flag_employee & ",'" & str_employee_code & "'," _
'            & cbo_yearly_division.ListIndex & ",'" & LOGIN_LEVEL & "');"
    str_sql = "SELECT a.company_code, e.company_name, a.division_code, f.division_name, " & _
                "a.employee_code, a.nik, a.employee_name, b.perf_avg, c.perf_grade, d.perf_date, " & _
                "d.perf_value, d.description, '" & IIf(LOGIN_LEVEL = 100, LOGIN_FULLNAME, EMPLOYEE_NAME) & "' " & _
              "FROM m_employee a join t_employee_perf_result b on a.employee_code = b.employee_code " & _
                "JOIN m_performance c on b.perf_number = c.perf_number " & _
                "JOIN t_employee_perf d on b.employee_code = d.employee_code AND b.perf_year = d.perf_year " & _
                "JOIN m_company e on a.company_code = e.company_code " & _
                "JOIN m_division f on a.company_code = f.company_code AND a.division_code = f.division_code " & _
              "WHERE " & IIf(cbo_yearly_division.ListIndex = 0, "a.company_code = '" & TDBCombo_company.Text & "' ", "a.company_code = '" & TDBCombo_company.Text & "' AND a.division_code = '" & TDBCombo_yearly_division.Text & "' ") & " " & _
                "" & IIf(cbo_yearly_employee.ListIndex = 0, "", "AND a.employee_code = '" & txt_yearly_employee_code.Text & " '") & " " & _
                "AND YEAR(d.perf_date) = '" & Format(DTPicker_yearly.Value, "yyyy") & "' " & _
              "ORDER BY a.employee_name"
    str_param_periode = Format(DTPicker_yearly.Value, "yyyy")
    
    Call a.Show
    a.Caption = "DETAIL PERFORMA KARYAWAN"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub cmdPrint_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
    
    If SSTab1.Tab = 0 Then
        If check_valid_yearly Then
            Call rpt_yearly
        End If
    End If
End Sub

Private Sub Form_Load()
    Adodc_company.ConnectionString = strConn
    Adodc_yearly_division.ConnectionString = strConn
    
    Call load_data_company
    
    DTPicker_yearly.Value = Now
    cbo_yearly_division.ListIndex = 0
    cbo_yearly_employee.ListIndex = 0
    
    Call createGridKar
    
    timer1.Enabled = True
    SSTab1.Tab = 0
End Sub

Private Sub load_data_company()
    Adodc_company.RecordSource = "select * from m_company order by company_code"
    Adodc_company.Refresh
    
    TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub TDBCombo_karyawan_ItemChange()
    If Not (TDBCombo_karyawan.ApproxCount > 0 And TDBCombo_karyawan.Bookmark > 0) Then Exit Sub
    
    TDBCombo_karyawan.Text = TDBCombo_karyawan.Columns("kode_karyawan").Value
    txt_nama_karyawan = TDBCombo_karyawan.Columns("nama_karyawan").Value
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        Call load_data_division_yearly
    End If
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If
    
    If SSTab1.Tab = 0 Then
        Call load_data_division_yearly
    End If
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode_adodc(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Private Sub load_data_division_yearly()
TDBCombo_yearly_division.Text = "": txt_yearly_division_name = ""

    Adodc_yearly_division.RecordSource = "select division_code, division_name from m_division where company_code='" _
    & TDBCombo_company.Columns("company_code").Value & "' order by division_code"
    Adodc_yearly_division.Refresh
    
    TDBCombo_yearly_division.RowSource = Adodc_yearly_division
End Sub

Private Sub TDBCombo_yearly_division_itemChange()
    If TDBCombo_yearly_division.ApproxCount > 0 Then
        TDBCombo_yearly_division.Text = TDBCombo_yearly_division.Columns("division_code").Value
        txt_yearly_division_name = TDBCombo_yearly_division.Columns("division_name").Value
    End If
End Sub


Private Sub createGridKar()
   With LynxGrid2
      .AddColumn "KODE KARY.", 1200, lgAlignCenterCenter, , , , , , , True
      .AddColumn "NAMA KARY.", 2000, , , , , , , , , True
      .AddColumn "Div. code", , , , , , , , , False
      .AddColumn "DIVISI", 1300, , , , , , , , , True
      .AddColumn "title code", , , , , , , , , False
      .AddColumn "JABATAN", 1300, , , , , , , , , True
      .AddColumn "Employee Code", 3000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear
        If LOGIN_LEVEL = 100 Then
            SQL = "SELECT a.nik,a.employee_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE " & IIf(cbo_yearly_division.ListIndex = 0, "b.company_code = '" & TDBCombo_company.Text & "'", _
                            "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_yearly_division.Text & "'") & " " _
                        & "AND (a.nik LIKE '%" & txt_yearly_nik.Text & "%' " _
                        & "OR a.employee_name LIKE '%" & txt_yearly_nik.Text & "%') " _
                        & "AND a.flag_active <> 0"
        Else
            SQL = "SELECT a.nik,a.employee_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE " & IIf(cbo_yearly_division.ListIndex = 0, "b.company_code = '" & TDBCombo_company.Text & "'", _
                            "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_yearly_division.Text & "'") & " " _
                        & "AND (a.nik LIKE '%" & txt_yearly_nik.Text & "%' " _
                        & "OR a.employee_name LIKE '%" & txt_yearly_nik.Text & "%') " _
                        & "AND a.flag_active <> 0 AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                        & "ORDER BY a.employee_name ASC"

        End If
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid2.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME _
                                & vbTab & rs!division_code & vbTab & rs!division_name _
                                & vbTab & rs!title_code & vbTab & rs!title_name _
                                & vbTab & rs!employee_code
                rs.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                txt_yearly_employee_code.Text = rs!employee_code
                txt_yearly_employee_name.Text = rs!EMPLOYEE_NAME
                txt_yearly_nik.Text = rs!nik
'                TDBCombo1.SetFocus
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else
            
        End If
        rs.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txt_yearly_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txt_yearly_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txt_yearly_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
        End If
        LynxGrid2.Visible = False
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

Private Sub txt_yearly_nik_Change()
    If txt_yearly_nik.Text = "" Then
        txt_yearly_employee_code.Text = ""
        txt_yearly_employee_name.Text = ""
    End If
End Sub

Private Sub txt_yearly_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmd_monthly_browse_employee_Click()
    isiGridKar (1)
End Sub

