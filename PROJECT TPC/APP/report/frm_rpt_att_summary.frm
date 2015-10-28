VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_rpt_summary_attendance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SUMMARY ATTENDANCE REPORT"
   ClientHeight    =   7095
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   10560
   Icon            =   "frm_rpt_att_summary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   17
      Top             =   780
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   1260
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "DAILY"
      TabPicture(0)   =   "frm_rpt_att_summary.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "MONTHLY"
      TabPicture(1)   =   "frm_rpt_att_summary.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PERIODE"
      TabPicture(2)   =   "frm_rpt_att_summary.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   12
         Top             =   960
         Width           =   8415
         Begin VB.TextBox txt_monthly_employee_code 
            Height          =   285
            Left            =   5070
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   540
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CheckBox chk_sort_monthly 
            Height          =   255
            Left            =   1800
            TabIndex        =   45
            Top             =   1920
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cbo_monthly_employee 
            Height          =   315
            ItemData        =   "frm_rpt_att_summary.frx":0060
            Left            =   1800
            List            =   "frm_rpt_att_summary.frx":006A
            TabIndex        =   30
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Frame fra_monthly_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   26
            Top             =   960
            Width           =   4575
            Begin VB.TextBox txt_monthly_nik 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   0
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   29
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txt_monthly_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   28
               Top             =   240
               Width           =   2415
            End
            Begin VB.CommandButton cmd_monthly_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1440
               TabIndex        =   27
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.ComboBox cbo_monthly_company 
            Height          =   315
            ItemData        =   "frm_rpt_att_summary.frx":007B
            Left            =   1800
            List            =   "frm_rpt_att_summary.frx":0085
            TabIndex        =   13
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_monthly 
            Height          =   300
            Left            =   1800
            TabIndex        =   14
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   110690307
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label15 
            Caption         =   "* yyyy-MM"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   3510
            TabIndex        =   54
            Top             =   1590
            Width           =   1065
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "ORDER"
            Height          =   195
            Left            =   720
            TabIndex        =   46
            Top             =   1920
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Karyawan"
            Height          =   195
            Left            =   720
            TabIndex        =   31
            Top             =   1200
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Bulan"
            Height          =   195
            Left            =   720
            TabIndex        =   16
            Top             =   1560
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Perusahaan"
            Height          =   195
            Left            =   720
            TabIndex        =   15
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   840
         TabIndex        =   4
         Top             =   960
         Width           =   8415
         Begin VB.TextBox txt_periode_employee_code 
            Height          =   285
            Left            =   4890
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   540
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.CheckBox chk_sort_periode 
            Height          =   255
            Left            =   1800
            TabIndex        =   43
            Top             =   1920
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Frame fra_periode_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   22
            Top             =   960
            Width           =   4575
            Begin VB.CommandButton cmd_periode_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1440
               TabIndex        =   25
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txt_periode_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   24
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox txt_periode_nik 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   0
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   23
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.ComboBox cbo_periode_employee 
            Height          =   315
            ItemData        =   "frm_rpt_att_summary.frx":0098
            Left            =   1800
            List            =   "frm_rpt_att_summary.frx":00A2
            TabIndex        =   20
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "DAY COUNT"
            Height          =   495
            Left            =   0
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cbo_periode_company 
            Height          =   315
            ItemData        =   "frm_rpt_att_summary.frx":00B3
            Left            =   1800
            List            =   "frm_rpt_att_summary.frx":00BD
            TabIndex        =   10
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox cbo_periode_to 
            Height          =   315
            ItemData        =   "frm_rpt_att_summary.frx":00D0
            Left            =   3600
            List            =   "frm_rpt_att_summary.frx":00DA
            TabIndex        =   7
            Text            =   "..."
            Top             =   1560
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_from 
            Height          =   300
            Left            =   1800
            TabIndex        =   5
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   232062979
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_to 
            Height          =   300
            Left            =   5040
            TabIndex        =   8
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   232062979
            CurrentDate     =   39278
         End
         Begin VB.Label Label16 
            Caption         =   "* yyyy-MM-dd"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   6780
            TabIndex        =   55
            Top             =   1590
            Width           =   1065
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "ORDER"
            Height          =   195
            Left            =   720
            TabIndex        =   44
            Top             =   1920
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Karyawan"
            Height          =   195
            Left            =   720
            TabIndex        =   21
            Top             =   1200
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Perusahaan"
            Height          =   195
            Left            =   720
            TabIndex        =   9
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Periode"
            Height          =   195
            Left            =   720
            TabIndex        =   6
            Top             =   1560
            Width           =   540
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   32
         Top             =   960
         Width           =   8415
         Begin VB.TextBox txt_daily_employee_code 
            Height          =   285
            Left            =   4650
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   570
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.CheckBox chk_sort_daily 
            Height          =   255
            Left            =   1800
            TabIndex        =   47
            Top             =   1920
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cbo_daily_company 
            Height          =   315
            ItemData        =   "frm_rpt_att_summary.frx":00E8
            Left            =   1800
            List            =   "frm_rpt_att_summary.frx":00F2
            TabIndex        =   38
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.Frame fra_daily_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   34
            Top             =   960
            Width           =   4575
            Begin VB.CommandButton cmd_daily_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1440
               TabIndex        =   37
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txt_daily_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   36
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox txt_daily_nik 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   0
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   35
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.ComboBox cbo_daily_employee 
            Height          =   315
            ItemData        =   "frm_rpt_att_summary.frx":0105
            Left            =   1800
            List            =   "frm_rpt_att_summary.frx":010F
            TabIndex        =   33
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_daily 
            Height          =   300
            Left            =   1800
            TabIndex        =   39
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   232390659
            CurrentDate     =   39278
         End
         Begin VB.Label Label14 
            Caption         =   "* yyyy-MM-dd"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   3600
            TabIndex        =   53
            Top             =   1620
            Width           =   1065
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "ORDER"
            Height          =   195
            Left            =   720
            TabIndex        =   48
            Top             =   1920
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Perusahaan"
            Height          =   195
            Left            =   720
            TabIndex        =   42
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Hari"
            Height          =   195
            Left            =   720
            TabIndex        =   41
            Top             =   1560
            Width           =   285
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Karyawan"
            Height          =   195
            Left            =   720
            TabIndex        =   40
            Top             =   1200
            Width           =   705
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   5700
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   840
         Top             =   360
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6480
         Picture         =   "frm_rpt_att_summary.frx":0120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   8280
         Picture         =   "frm_rpt_att_summary.frx":06AA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_rpt_att_summary.frx":0C34
      TabIndex        =   18
      Top             =   780
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1080
      Top             =   780
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
      Top             =   780
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
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "LAPORAN REKAP KEHADIRAN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3300
      TabIndex        =   52
      Top             =   0
      Width           =   4785
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   780
      Width           =   795
   End
End
Attribute VB_Name = "frm_rpt_summary_attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim int_libur() As Integer


Private Sub report_kehadiran_karyawan()
Dim lstr_criteria As String

CrystalReport1.Reset
CrystalReport1.ReportFileName = App.Path & "\report\rpt_kehadiran_karyawan.rpt"
    
lstr_criteria = "({v_att_karyawan.flag_io}) = 0 and Date({v_att_karyawan.tanggal}) in #" _
                & Format(DTPicker_bulan.Value, "yyyy,mm,01") & "# to #" _
                & Format(DTPicker_bulan.Value, "yyyy,mm,31") & "#"
    
CrystalReport1.ParameterFields(0) = "p_karyawan;Karyawan : (" _
& TDBCombo_karyawan.Columns("kode_karyawan").Value & ") " _
& TDBCombo_karyawan.Columns("nama_karyawan").Value & ";true"

CrystalReport1.ReplaceSelectionFormula (lstr_criteria)

CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
End Sub

Private Sub cbo_daily_company_Click()
If cbo_daily_company.ListIndex = 0 Then
    cbo_daily_employee.ListIndex = 0
    cbo_daily_employee.Enabled = False
ElseIf cbo_daily_company.ListIndex = 1 Then
    cbo_daily_employee.ListIndex = 0
    cbo_daily_employee.Enabled = True
End If
End Sub

Private Sub cbo_daily_employee_Click()
If cbo_daily_employee.ListIndex = 0 Then
    fra_daily_employee.Visible = False
Else
    fra_daily_employee.Visible = True
End If

txt_daily_employee_code = "": txt_daily_employee_name = "": txt_daily_nik = ""
End Sub

Private Sub cbo_monthly_company_Click()
If cbo_monthly_company.ListIndex = 0 Then
    cbo_monthly_employee.ListIndex = 0
    cbo_monthly_employee.Enabled = False
ElseIf cbo_monthly_company.ListIndex = 1 Then
    cbo_monthly_employee.ListIndex = 0
    cbo_monthly_employee.Enabled = True
End If
End Sub

Private Sub cbo_monthly_employee_Click()
If cbo_monthly_employee.ListIndex = 0 Then
    fra_monthly_employee.Visible = False
Else
    fra_monthly_employee.Visible = True
End If

txt_monthly_employee_code = "": txt_monthly_employee_name = "": txt_monthly_nik = ""
End Sub

Private Sub cbo_periode_company_Click()
If cbo_periode_company.ListIndex = 0 Then
    cbo_periode_employee.ListIndex = 0
    cbo_periode_employee.Enabled = False
ElseIf cbo_periode_company.ListIndex = 1 Then
    cbo_periode_employee.ListIndex = 0
    cbo_periode_employee.Enabled = True
End If
End Sub

Private Sub cbo_periode_employee_Click()
If cbo_periode_employee.ListIndex = 0 Then
    fra_periode_employee.Visible = False
Else
    fra_periode_employee.Visible = True
End If

txt_periode_employee_code = "": txt_periode_employee_name = "": txt_periode_nik = ""
End Sub

Private Sub cmd_daily_browse_employee_Click()
frm_lookup_mst_employee.public_int_mode = 9
frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_monthly_browse_employee_Click()
frm_lookup_mst_employee.public_int_mode = 8
frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_periode_browse_employee_Click()
frm_lookup_mst_employee.public_int_mode = 7
frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
frm_lookup_mst_employee.Show 1
End Sub

Private Sub Command1_Click()
'MsgBox TDBDate1.MinDate & " - " & TDBDate1.MaxDate _
& vbCr & TDBDate1.Month _
& vbCr & TDBDate1.Weekday


'MsgBox DTPicker1.DayOfWeek
'ReDim int_libur(5)
Dim a, b As Date

'a = DTPicker_periode_from.Value
'b = DateAdd("m", 1, a)

a = Format(DTPicker_periode_from.Value, "yyyy-MM-01")
b = DateAdd("m", 1, a)
'DTPicker_periode_to.Value = DateAdd("m", 1, DTPicker_periode_from.Value)

'MsgBox DTPicker_periode_from.Value & vbCr _
& DateDiff("d", DTPicker_periode_from.Value, DTPicker_periode_to.Value)

MsgBox a & vbCr _
& DateDiff("d", a, b)
End Sub

Private Sub Command2_Click()
'MsgBox UBound(int_libur)

'Call get_holiday("2007-07")

Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

cmd.ActiveConnection = CnG
cmd.CommandText = Text1.Text
rs.CursorLocation = adUseClient
rs.Open cmd, , adOpenStatic, adLockReadOnly

MsgBox rs.RecordCount
End Sub


Private Sub get_holiday(ByVal i As String)
DTPicker1.Value = i & "-01"
End Sub

Private Sub periode_date_event()
If cbo_periode_to.ListIndex = 0 Then
    DTPicker_periode_to.Visible = False
Else
    DTPicker_periode_to.Visible = True
    DTPicker_periode_to.Value = DTPicker_periode_from.Value
End If
End Sub

Private Sub cbo_periode_to_Click()
Call periode_date_event
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Function check_valid_periode() As Boolean
check_valid_periode = True

'validate employee
If cbo_periode_employee.ListIndex = 1 And Trim(txt_periode_employee_code) = "" Then
    MsgBox "Employee is not selected!", vbOKOnly + vbInformation, headerMSG
    cmd_periode_browse_employee.SetFocus
    check_valid_periode = False
    Exit Function
End If
End Function

Private Function check_valid_monthly() As Boolean
check_valid_monthly = True

'validate employee
If cbo_monthly_employee.ListIndex = 1 And Trim(txt_monthly_employee_code) = "" Then
    MsgBox "Employee is not selected!", vbOKOnly + vbInformation, headerMSG
    cmd_monthly_browse_employee.SetFocus
    check_valid_monthly = False
    Exit Function
End If
End Function

Private Function check_valid_daily() As Boolean
check_valid_daily = True

'validate employee
If cbo_daily_employee.ListIndex = 1 And Trim(txt_daily_employee_code) = "" Then
    MsgBox "Employee is not selected!", vbOKOnly + vbInformation, headerMSG
    cmd_daily_browse_employee.SetFocus
    check_valid_daily = False
    Exit Function
End If
End Function

'Private Sub rpt_periode()
'Dim str_sql, str_param_periode, str_file As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim a As New frm_rpt
'
'If cbo_periode_company.ListIndex = 0 Then
'    int_flag_company = 0
'    str_company_code = "-"
'ElseIf cbo_periode_company.ListIndex = 1 Then
'    int_flag_company = 1
'    str_company_code = TDBCombo_company.Columns("company_code").Value
'End If
'
'str_file = "\report\rpt_11_mdn.rpt"
'    If chk_sort_periode Then
'        str_file = "\report\rpt_11_mdn_sort.rpt"
'    End If
'
''If mdi_absensi.mnu_stg_language_english.Checked Then
''    str_file = "\report\rpt_11_mdn.rpt"
''    If chk_sort_periode Then
''        str_file = "\report\rpt_11_mdn_sort.rpt"
''    End If
''ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
''    str_file = "\report\rpt_11_ind.rpt"
''    If chk_sort_periode Then
''        str_file = "\report\rpt_11_ind_sort.rpt"
''    End If
''End If
'
'If int_flag_company = 0 Then
'    int_flag_employee = 0
'    str_employee_code = "-"
'ElseIf int_flag_company = 1 Then
'    If cbo_periode_employee.ListIndex = 0 Then
'        int_flag_employee = 0
'        str_employee_code = "-"
'    ElseIf cbo_periode_employee.ListIndex = 1 Then
'        int_flag_employee = 1
'        str_employee_code = txt_periode_employee_code
'    End If
'End If
'
'If cbo_periode_to.ListIndex = 0 Then
'    str_sql = "call spr_attendance_12('" _
'        & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" _
'        & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "'," _
'        & int_flag_company & ",'" & str_company_code & "'," _
'        & int_flag_employee & ",'" & str_employee_code & "')"
'    str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
'ElseIf cbo_periode_to.ListIndex = 1 Then
'    str_sql = "call spr_attendance_12('" _
'        & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" _
'        & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "'," _
'        & int_flag_company & ",'" & str_company_code & "'," _
'        & int_flag_employee & ",'" & str_employee_code & "')"
'    str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
'        & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
'End If
'
'Call a.Show
'a.Caption = "SUMMARY REPORT"
''If mdi_absensi.mnu_stg_language_english.Checked Then
''    a.Caption = "SUMMARY REPORT"
''ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
''    a.Caption = "LAPORAN REKAP"
''End If
'Call a.rpt_view(str_sql, str_file, str_param_periode)
'End Sub

Private Sub rpt_periode()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt

If cbo_periode_company.ListIndex = 0 Then
    int_flag_company = 0
    str_company_code = "-"
ElseIf cbo_periode_company.ListIndex = 1 Then
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
End If

'str_file = "\report\rpt_11_mdn.rpt"
str_file = "\report\rpt_summary_attendance.rpt"
    If chk_sort_periode Then
'        str_file = "\report\rpt_11_mdn_sort.rpt"
        str_file = "\report\rpt_summary_attendance_sort.rpt"
    End If

'If mdi_absensi.mnu_stg_language_english.Checked Then
'    str_file = "\report\rpt_11_mdn.rpt"
'    If chk_sort_periode Then
'        str_file = "\report\rpt_11_mdn_sort.rpt"
'    End If
'ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    str_file = "\report\rpt_11_ind.rpt"
'    If chk_sort_periode Then
'        str_file = "\report\rpt_11_ind_sort.rpt"
'    End If
'End If

If int_flag_company = 0 Then
    int_flag_employee = 0
    str_employee_code = "-"
ElseIf int_flag_company = 1 Then
    If cbo_periode_employee.ListIndex = 0 Then
        int_flag_employee = 0
        str_employee_code = "-"
    ElseIf cbo_periode_employee.ListIndex = 1 Then
        int_flag_employee = 1
        str_employee_code = txt_periode_employee_code
    End If
End If

If cbo_periode_to.ListIndex = 0 Then
    str_sql = "call spr_summary_attendance('" _
        & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" _
        & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "','" & LOGIN_CODE & "')"
    str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
ElseIf cbo_periode_to.ListIndex = 1 Then
    str_sql = "call spr_summary_attendance('" _
        & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" _
        & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "','" & LOGIN_CODE & "')"
    str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
        & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
End If

Call a.Show
a.Caption = "SUMMARY REPORT"
'If mdi_absensi.mnu_stg_language_english.Checked Then
'    a.Caption = "SUMMARY REPORT"
'ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    a.Caption = "LAPORAN REKAP"
'End If
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

'Private Sub rpt_daily()
'Dim str_sql, str_param_periode, str_file As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim a As New frm_rpt
'
'If cbo_daily_company.ListIndex = 0 Then
'    int_flag_company = 0
'    str_company_code = "-"
'ElseIf cbo_daily_company.ListIndex = 1 Then
'    int_flag_company = 1
'    str_company_code = TDBCombo_company.Columns("company_code").Value
'End If
'
'str_file = "\report\rpt_11_mdn.rpt"
'    If chk_sort_daily Then
'        str_file = "\report\rpt_11_mdn_sort.rpt"
'    End If
'
''If mdi_absensi.mnu_stg_language_english.Checked Then
''    str_file = "\report\rpt_11_mdn.rpt"
''    If chk_sort_daily Then
''        str_file = "\report\rpt_11_mdn_sort.rpt"
''    End If
''ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
''    str_file = "\report\rpt_11_ind.rpt"
''    If chk_sort_daily Then
''        str_file = "\report\rpt_11_ind_sort.rpt"
''    End If
''End If
'
'If int_flag_company = 0 Then
'    int_flag_employee = 0
'    str_employee_code = "-"
'ElseIf int_flag_company = 1 Then
'    If cbo_daily_employee.ListIndex = 0 Then
'        int_flag_employee = 0
'        str_employee_code = "-"
'    ElseIf cbo_daily_employee.ListIndex = 1 Then
'        int_flag_employee = 1
'        str_employee_code = txt_daily_employee_code
'    End If
'End If
'
'str_sql = "call spr_attendance_12('" _
'    & Format(DTPicker_daily.Value, "yyyy-MM-dd") & "','" _
'    & Format(DTPicker_daily.Value, "yyyy-MM-dd") & "'," _
'    & int_flag_company & ",'" & str_company_code & "'," _
'    & int_flag_employee & ",'" & str_employee_code & "')"
'
'Call a.Show
'a.Caption = "SUMMARY REPORT"
'str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
'
''If mdi_absensi.mnu_stg_language_english.Checked Then
''    a.Caption = "SUMMARY REPORT"
''    str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
''ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
''    a.Caption = "LAPORAN REKAP"
''    str_param_periode = "TGL : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
''End If
'Call a.rpt_view(str_sql, str_file, str_param_periode)
'End Sub

Private Sub rpt_daily()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt

If cbo_daily_company.ListIndex = 0 Then
    int_flag_company = 0
    str_company_code = "-"
ElseIf cbo_daily_company.ListIndex = 1 Then
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
End If

'str_file = "\report\rpt_11_mdn.rpt"
str_file = "\report\rpt_summary_attendance.rpt"
    If chk_sort_daily Then
'        str_file = "\report\rpt_11_mdn_sort.rpt"
        str_file = "\report\rpt_summary_attendance_sort.rpt"
    End If
    
'If mdi_absensi.mnu_stg_language_english.Checked Then
'    str_file = "\report\rpt_11_mdn.rpt"
'    If chk_sort_daily Then
'        str_file = "\report\rpt_11_mdn_sort.rpt"
'    End If
'ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    str_file = "\report\rpt_11_ind.rpt"
'    If chk_sort_daily Then
'        str_file = "\report\rpt_11_ind_sort.rpt"
'    End If
'End If

If int_flag_company = 0 Then
    int_flag_employee = 0
    str_employee_code = "-"
ElseIf int_flag_company = 1 Then
    If cbo_daily_employee.ListIndex = 0 Then
        int_flag_employee = 0
        str_employee_code = "-"
    ElseIf cbo_daily_employee.ListIndex = 1 Then
        int_flag_employee = 1
        str_employee_code = txt_daily_employee_code
    End If
End If

str_sql = "call spr_summary_attendance('" _
    & Format(DTPicker_daily.Value, "yyyy-MM-dd") & "','" _
    & Format(DTPicker_daily.Value, "yyyy-MM-dd") & "'," _
    & int_flag_company & ",'" & str_company_code & "'," _
    & int_flag_employee & ",'" & str_employee_code & "','" & LOGIN_CODE & "')"

Call a.Show
a.Caption = "SUMMARY REPORT"
str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
 
'If mdi_absensi.mnu_stg_language_english.Checked Then
'    a.Caption = "SUMMARY REPORT"
'    str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
'ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    a.Caption = "LAPORAN REKAP"
'    str_param_periode = "TGL : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
'End If
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

'Private Sub rpt_monthly()
'Dim str_sql, str_param_periode, str_file As String
'Dim int_flag_company As Integer, str_company_code As String
'Dim int_flag_employee As Integer, str_employee_code As String
'Dim a As New frm_rpt
'Dim d1, d2 As Date
'
'If cbo_monthly_company.ListIndex = 0 Then
'    int_flag_company = 0
'    str_company_code = "-"
'ElseIf cbo_monthly_company.ListIndex = 1 Then
'    int_flag_company = 1
'    str_company_code = TDBCombo_company.Columns("company_code").Value
'End If
'
'str_file = "\report\rpt_11_mdn.rpt"
'    If chk_sort_monthly Then
'        str_file = "\report\rpt_11_mdn_sort.rpt"
'    End If
'
''If mdi_absensi.mnu_stg_language_english.Checked Then
''    str_file = "\report\rpt_11_mdn.rpt"
''    If chk_sort_monthly Then
''        str_file = "\report\rpt_11_mdn_sort.rpt"
''    End If
''ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
''    str_file = "\report\rpt_11_ind.rpt"
''    If chk_sort_monthly Then
''        str_file = "\report\rpt_11_ind_sort.rpt"
''    End If
''End If
'
'If int_flag_company = 0 Then
'    int_flag_employee = 0
'    str_employee_code = "-"
'ElseIf int_flag_company = 1 Then
'    If cbo_monthly_employee.ListIndex = 0 Then
'        int_flag_employee = 0
'        str_employee_code = "-"
'    ElseIf cbo_monthly_employee.ListIndex = 1 Then
'        int_flag_employee = 1
'        str_employee_code = txt_monthly_employee_code
'    End If
'End If
'
'd1 = Format(DTPicker_monthly.Value, "yyyy-MM-01"): d2 = DateAdd("m", 1, d1)
'str_sql = "call spr_attendance_12('" & d1 & "','" _
'    & Format(d1, "yyyy-MM-") & Format(DateDiff("d", d1, d2), "0#") & "'," _
'    & int_flag_company & ",'" & str_company_code & "'," _
'    & int_flag_employee & ",'" & str_employee_code & "')"
'
'
'Call a.Show
'a.Caption = "SUMMARY REPORT"
'str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
'
''If mdi_absensi.mnu_stg_language_english.Checked Then
''    a.Caption = "SUMMARY REPORT"
''    str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
''ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
''    a.Caption = "LAPORAN REKAP"
''    str_param_periode = "BULAN : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
''End If
'
'Call a.rpt_view(str_sql, str_file, str_param_periode)
'
'End Sub

Private Sub rpt_monthly()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2 As Date

If cbo_monthly_company.ListIndex = 0 Then
    int_flag_company = 0
    str_company_code = "-"
ElseIf cbo_monthly_company.ListIndex = 1 Then
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
End If

'str_file = "\report\rpt_11_mdn.rpt"
str_file = "\report\rpt_summary_attendance.rpt"
    If chk_sort_monthly Then
'        str_file = "\report\rpt_11_mdn_sort.rpt"
        str_file = "\report\rpt_summary_attendance_sort.rpt"
    End If
    
'If mdi_absensi.mnu_stg_language_english.Checked Then
'    str_file = "\report\rpt_11_mdn.rpt"
'    If chk_sort_monthly Then
'        str_file = "\report\rpt_11_mdn_sort.rpt"
'    End If
'ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    str_file = "\report\rpt_11_ind.rpt"
'    If chk_sort_monthly Then
'        str_file = "\report\rpt_11_ind_sort.rpt"
'    End If
'End If

If int_flag_company = 0 Then
    int_flag_employee = 0
    str_employee_code = "-"
ElseIf int_flag_company = 1 Then
    If cbo_monthly_employee.ListIndex = 0 Then
        int_flag_employee = 0
        str_employee_code = "-"
    ElseIf cbo_monthly_employee.ListIndex = 1 Then
        int_flag_employee = 1
        str_employee_code = txt_monthly_employee_code
    End If
End If

d1 = Format(DTPicker_monthly.Value, "yyyy-MM-01"): d2 = DateAdd("m", 1, d1)
str_sql = "call spr_summary_attendance('" & d1 & "','" _
    & Format(d1, "yyyy-MM-") & Format(DateDiff("d", d1, d2), "0#") & "'," _
    & int_flag_company & ",'" & str_company_code & "'," _
    & int_flag_employee & ",'" & str_employee_code & "','" & LOGIN_CODE & "')"


Call a.Show
a.Caption = "SUMMARY REPORT"
str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"

'If mdi_absensi.mnu_stg_language_english.Checked Then
'    a.Caption = "SUMMARY REPORT"
'    str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
'ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    a.Caption = "LAPORAN REKAP"
'    str_param_periode = "BULAN : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
'End If

Call a.rpt_view(str_sql, str_file, str_param_periode)

End Sub

Private Sub cmdPrint_Click()
If check_validate_tdbcombo(TDBCombo_company) = False Then
    MsgBox "No Company selected!", vbInformation, headerMSG
    Exit Sub
End If

If SSTab1.Tab = 2 Then
    If check_valid_periode Then
        Call rpt_periode
    End If
ElseIf SSTab1.Tab = 1 Then
    If check_valid_monthly Then
        Call rpt_monthly
    End If
ElseIf SSTab1.Tab = 0 Then
    If check_valid_daily Then
        Call rpt_daily
    End If
End If
End Sub

Private Sub Command3_Click()
Adodc1.ConnectionString = strConn
Adodc1.RecordSource = Text1.Text
Adodc1.Refresh

MsgBox Adodc1.Recordset.RecordCount
End Sub

Private Sub Form_Load()
Adodc_company.ConnectionString = strConn

Call load_data_company

DTPicker_periode_from.Value = Now
DTPicker_periode_to.Value = Now
DTPicker_monthly.Value = Now
DTPicker_daily.Value = Now
Call load_data_user_access(Me)

cbo_periode_to.ListIndex = 0
cbo_periode_company.ListIndex = 1
cbo_periode_employee.ListIndex = 0

cbo_monthly_company.ListIndex = 1
cbo_monthly_employee.ListIndex = 0

cbo_daily_company.ListIndex = 1
cbo_daily_employee.ListIndex = 0

Timer1.Enabled = True
SSTab1.Tab = 0
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub load_data_monthly_company()
Adodc_monthly_company.RecordSource = "select * from m_company order by company_code"
Adodc_monthly_company.Refresh

TDBCombo_monthly_company.RowSource = Adodc_monthly_company
End Sub

Private Sub TDBCombo_karyawan_ItemChange()
If Not (TDBCombo_karyawan.ApproxCount > 0 And TDBCombo_karyawan.Bookmark > 0) Then Exit Sub

TDBCombo_karyawan.Text = TDBCombo_karyawan.Columns("kode_karyawan").Value
txt_nama_karyawan = TDBCombo_karyawan.Columns("nama_karyawan").Value
End Sub

Private Sub set_company_option()
If opt_per_company Then
    TDBGrid1.Enabled = True
ElseIf opt_all Then
    TDBGrid1.Enabled = False
End If
End Sub

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
End If
End Sub

Private Sub TDBCombo_monthly_company_ItemChange()
If TDBCombo_monthly_company.ApproxCount > 0 Then
    TDBCombo_monthly_company.Text = TDBCombo_monthly_company.Columns("company_code").Value
    txt_monthly_company_name = TDBCombo_monthly_company.Columns("company_name").Value
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
MsgBox KeyAscii
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
If LOGIN_LEVEL = 100 Then
    cbo_daily_company.Enabled = True
    cbo_monthly_company.Enabled = True
    cbo_periode_company.Enabled = True
Else
    cbo_daily_company.Enabled = False
    cbo_monthly_company.Enabled = False
    cbo_periode_company.Enabled = False
End If
End Sub
