VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_rpt_att_detail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DETAIL ATTENDANCE REPORT"
   ClientHeight    =   6570
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   10560
   Icon            =   "frm_rpt_att_detail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin prj_absensi.LynxGrid LynxGrid1 
      Height          =   2805
      Left            =   4680
      TabIndex        =   42
      Top             =   3210
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4948
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
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   20
      Top             =   240
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "DAILY"
      TabPicture(0)   =   "frm_rpt_att_detail.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "MONTHLY"
      TabPicture(1)   =   "frm_rpt_att_detail.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PERIODE"
      TabPicture(2)   =   "frm_rpt_att_detail.frx":05C2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   17
         Top             =   960
         Width           =   8415
         Begin VB.ComboBox cbo_monthly_employee 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":05DE
            Left            =   1800
            List            =   "frm_rpt_att_detail.frx":05E8
            TabIndex        =   4
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Frame fra_monthly_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   27
            Top             =   960
            Width           =   4575
            Begin VB.TextBox txt_monthly_employee_code 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   0
               MaxLength       =   50
               TabIndex        =   29
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txt_monthly_employee_name 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   28
               Top             =   240
               Width           =   2415
            End
            Begin prj_absensi.vbButton cmd_monthly_browse_employee 
               Height          =   315
               Left            =   1440
               TabIndex        =   40
               Top             =   240
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   556
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
               MICON           =   "frm_rpt_att_detail.frx":05F9
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
         Begin VB.ComboBox cbo_monthly_company 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":0615
            Left            =   1800
            List            =   "frm_rpt_att_detail.frx":061F
            TabIndex        =   3
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_monthly 
            Height          =   300
            Left            =   1800
            TabIndex        =   5
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   96141315
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   720
            TabIndex        =   30
            Top             =   1200
            Width           =   870
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "MONTH"
            Height          =   195
            Left            =   720
            TabIndex        =   19
            Top             =   1560
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "COMPANY"
            Height          =   195
            Left            =   720
            TabIndex        =   18
            Top             =   840
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   840
         TabIndex        =   13
         Top             =   960
         Width           =   8415
         Begin VB.Frame fra_periode_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   24
            Top             =   960
            Width           =   4575
            Begin VB.TextBox txt_periode_employee_name 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   26
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox txt_periode_employee_code 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   0
               MaxLength       =   50
               TabIndex        =   25
               Top             =   240
               Width           =   1335
            End
            Begin prj_absensi.vbButton cmd_periode_browse_employee 
               Height          =   315
               Left            =   1440
               TabIndex        =   41
               Top             =   240
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   556
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
               MICON           =   "frm_rpt_att_detail.frx":0632
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
         Begin VB.ComboBox cbo_periode_employee 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":064E
            Left            =   1800
            List            =   "frm_rpt_att_detail.frx":0658
            TabIndex        =   7
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "DAY COUNT"
            Height          =   495
            Left            =   0
            TabIndex        =   16
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cbo_periode_company 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":0669
            Left            =   1800
            List            =   "frm_rpt_att_detail.frx":0673
            TabIndex        =   6
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox cbo_periode_to 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":0686
            Left            =   3600
            List            =   "frm_rpt_att_detail.frx":0690
            TabIndex        =   9
            Text            =   "..."
            Top             =   1560
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_from 
            Height          =   300
            Left            =   1800
            TabIndex        =   8
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_to 
            Height          =   300
            Left            =   5040
            TabIndex        =   10
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39278
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   720
            TabIndex        =   23
            Top             =   1200
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "COMPANY"
            Height          =   195
            Left            =   720
            TabIndex        =   15
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PERIODE"
            Height          =   195
            Left            =   720
            TabIndex        =   14
            Top             =   1560
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   31
         Top             =   960
         Width           =   8415
         Begin VB.ComboBox cbo_daily_company 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":069E
            Left            =   1800
            List            =   "frm_rpt_att_detail.frx":06A8
            TabIndex        =   0
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.Frame fra_daily_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   32
            Top             =   960
            Width           =   4575
            Begin VB.TextBox txt_daily_employee_name 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   34
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox txt_daily_employee_code 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   0
               MaxLength       =   50
               TabIndex        =   33
               Top             =   240
               Width           =   1335
            End
            Begin prj_absensi.vbButton cmd_daily_browse_employee 
               Height          =   315
               Left            =   1440
               TabIndex        =   43
               Top             =   240
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   556
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
               MICON           =   "frm_rpt_att_detail.frx":06BB
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
         Begin VB.ComboBox cbo_daily_employee 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":06D7
            Left            =   1800
            List            =   "frm_rpt_att_detail.frx":06E1
            TabIndex        =   1
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_daily 
            Height          =   300
            Left            =   1800
            TabIndex        =   2
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96141315
            CurrentDate     =   39278
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "COMPANY"
            Height          =   195
            Left            =   720
            TabIndex        =   37
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "DAY"
            Height          =   195
            Left            =   720
            TabIndex        =   36
            Top             =   1560
            Width           =   330
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   720
            TabIndex        =   35
            Top             =   1200
            Width           =   870
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   11
      Top             =   5160
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   840
         Top             =   360
      End
      Begin prj_absensi.vbButton cmdExit 
         Height          =   705
         Left            =   8220
         TabIndex        =   38
         Top             =   330
         Width           =   945
         _ExtentX        =   1667
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
         MICON           =   "frm_rpt_att_detail.frx":06F2
         PICN            =   "frm_rpt_att_detail.frx":070E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_absensi.vbButton cmdPrint 
         Height          =   705
         Left            =   7200
         TabIndex        =   39
         Top             =   330
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "Print"
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
         MICON           =   "frm_rpt_att_detail.frx":17A0
         PICN            =   "frm_rpt_att_detail.frx":17BC
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_rpt_att_detail.frx":284E
      TabIndex        =   21
      Top             =   240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1080
      Top             =   240
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
      Top             =   240
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frm_rpt_att_detail"
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
    txt_daily_employee_code = "": txt_daily_employee_name = ""
End If
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
    txt_monthly_employee_code = "": txt_monthly_employee_name = ""
End If
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
    txt_periode_employee_code = "": txt_periode_employee_name = ""
End If
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

Private Sub rpt_periode()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt

str_file = "\report\rpt_01.rpt"
If cbo_periode_company.ListIndex = 0 Then
    int_flag_company = 0
    str_company_code = "-"
ElseIf cbo_periode_company.ListIndex = 1 Then
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
End If

'If mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    str_file = "\report\rpt_01_ind.rpt"
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
    str_sql = "call spr_attendance_05('" _
        & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" _
        & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & ")"
ElseIf cbo_periode_to.ListIndex = 1 Then
    str_sql = "call spr_attendance_05('" _
        & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" _
        & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "'," _
        & int_flag_company & ",'" & str_company_code & "'," _
        & int_flag_employee & ",'" & str_employee_code & "')"
    str_param_periode = "PERIODE : (" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") _
        & " to " & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & ")"
End If

Call a.Show
'If mdi_absensi.mnu_stg_language_english.Checked Then
'    a.Caption = "DETAIL REPORT"
'ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    a.Caption = "LAPORAN RINCI"
'End If

a.Caption = "DETAIL REPORT"
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub rpt_daily()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt

str_file = "\report\rpt_01_1.rpt"
If cbo_daily_company.ListIndex = 0 Then
    int_flag_company = 0
    str_company_code = "-"
ElseIf cbo_daily_company.ListIndex = 1 Then
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
End If

'If mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    str_file = "\report\rpt_01_1_ind.rpt"
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

str_sql = "call spr_attendance_05('" _
    & Format(DTPicker_daily.Value, "yyyy-MM-dd") & "','" _
    & Format(DTPicker_daily.Value, "yyyy-MM-dd") & "'," _
    & int_flag_company & ",'" & str_company_code & "'," _
    & int_flag_employee & ",'" & str_employee_code & "')"


Call a.Show
'If mdi_absensi.mnu_stg_language_english.Checked Then
'    a.Caption = "DETAIL REPORT"
'    str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
'ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    a.Caption = "LAPORAN RINCI"
'    str_param_periode = "TGL : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
'End If

a.Caption = "DETAIL REPORT"
str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub rpt_monthly()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2 As Date

str_file = "\report\rpt_01.rpt"
If cbo_monthly_company.ListIndex = 0 Then
    int_flag_company = 0
    str_company_code = "-"
ElseIf cbo_monthly_company.ListIndex = 1 Then
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
End If

'If mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    str_file = "\report\rpt_01_ind.rpt"
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
str_sql = "call spr_attendance_05('" & d1 & "','" _
    & Format(d1, "yyyy-MM-") & Format(DateDiff("d", d1, d2), "0#") & "'," _
    & int_flag_company & ",'" & str_company_code & "'," _
    & int_flag_employee & ",'" & str_employee_code & "')"

Call a.Show
'If mdi_absensi.mnu_stg_language_english.Checked Then
'    a.Caption = "DETAIL REPORT"
'    str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
'ElseIf mdi_absensi.mnu_stg_language_indonesia.Checked Then
'    a.Caption = "LAPORAN RINCI"
'    str_param_periode = "BULAN : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
'End If

a.Caption = "DETAIL REPORT"
str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub CmdPrint_Click()
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
Call createGridKar

DTPicker_periode_from.Value = Now
DTPicker_periode_to.Value = Now
DTPicker_monthly.Value = Now
DTPicker_daily.Value = Now

cbo_periode_to.ListIndex = 0
cbo_periode_company.ListIndex = 1
cbo_periode_employee.ListIndex = 0

cbo_monthly_company.ListIndex = 1
cbo_monthly_employee.ListIndex = 0

cbo_daily_company.ListIndex = 1
cbo_daily_employee.ListIndex = 0

timer1.Enabled = True
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
timer1.Enabled = False
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

Private Sub createGridKar()
   With LynxGrid1
      .AddColumn "CODE", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "NAME", 2000, , , , , , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
Dim vEmployeeCode As String
    If pilihan = 1 Then
        LynxGrid1.Clear
        
        vEmployeeCode = IIf(SSTab1.Tab = 0, txt_daily_employee_code.Text, IIf(SSTab1.Tab = 1, txt_monthly_employee_code.Text, txt_periode_employee_code.Text))
        If rs.State Then rs.Close
        SQL = "SELECT employee_code, employee_name " & _
              "FROM m_employee " & _
                 "WHERE flag_active <> 0 " & _
                    "AND company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                    "AND (employee_code LIKE '%" & vEmployeeCode & "%' " & _
                        "OR employee_name LIKE '%" & vEmployeeCode & "%')"
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid1.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid1.AddItem rs!employee_code & vbTab & rs!EMPLOYEE_NAME
                rs.MoveNext
            Wend
            LynxGrid1.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                
                If SSTab1.Tab = 0 Then
                    txt_daily_employee_code.Text = rs!employee_code
                    txt_daily_employee_name.Text = rs!EMPLOYEE_NAME
                ElseIf SSTab1.Tab = 1 Then
                    txt_monthly_employee_code.Text = rs!employee_code
                    txt_monthly_employee_name.Text = rs!EMPLOYEE_NAME
                ElseIf SSTab1.Tab = 2 Then
                    txt_periode_employee_code.Text = rs!employee_code
                    txt_periode_employee_name.Text = rs!EMPLOYEE_NAME
                End If
            Else
                LynxGrid1.Visible = True
                LynxGrid1.SetFocus
            End If
        Else
            
        End If
        rs.Close
    Else
        If LynxGrid1.Rows > 0 Then
            If SSTab1.Tab = 0 Then
                txt_daily_employee_code.Text = LynxGrid1.CellText(LynxGrid1.Row, 0)
                txt_daily_employee_name.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
            ElseIf SSTab1.Tab = 1 Then
                txt_monthly_employee_code.Text = LynxGrid1.CellText(LynxGrid1.Row, 0)
                txt_monthly_employee_name.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
            ElseIf SSTab1.Tab = 2 Then
                txt_periode_employee_code.Text = LynxGrid1.CellText(LynxGrid1.Row, 0)
                txt_periode_employee_name.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
            End If
        End If
        LynxGrid1.Visible = False
    End If
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

Private Sub txt_daily_employee_code_Change()
    If txt_daily_employee_code.Text = "" Then
        txt_daily_employee_code.Text = ""
        txt_daily_employee_name.Text = ""
    End If
End Sub

Private Sub txt_monthly_employee_code_Change()
    If txt_monthly_employee_code.Text = "" Then
        txt_monthly_employee_code.Text = ""
        txt_monthly_employee_name.Text = ""
    End If
End Sub

Private Sub txt_periode_employee_code_Change()
    If txt_periode_employee_code.Text = "" Then
        txt_periode_employee_code.Text = ""
        txt_periode_employee_name.Text = ""
    End If
End Sub

Private Sub txt_daily_employee_code_KeyPress(KeyAscii As Integer)
    If TDBCombo_company.Text = "" Then
        MsgBox "Company not selected...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub txt_monthly_employee_code_KeyPress(KeyAscii As Integer)
    If TDBCombo_company.Text = "" Then
        MsgBox "Company not selected...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub txt_periode_employee_code_KeyPress(KeyAscii As Integer)
    If TDBCombo_company.Text = "" Then
        MsgBox "Company not selected...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmd_daily_browse_employee_Click()
    If TDBCombo_company.Text = "" Then
        MsgBox "Company not selected...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    isiGridKar (1)
End Sub

Private Sub cmd_monthly_browse_employee_Click()
    If TDBCombo_company.Text = "" Then
        MsgBox "Company not selected...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    isiGridKar (1)
End Sub

Private Sub cmd_periode_browse_employee_Click()
    If TDBCombo_company.Text = "" Then
        MsgBox "Company not selected...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    isiGridKar (1)
End Sub
