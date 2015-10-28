VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_rpt_layout_absen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "REPORT PRINT OUT ATTENDANCE"
   ClientHeight    =   7920
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   10560
   Icon            =   "frm_rpt_layout_absen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   780
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   240
      TabIndex        =   2
      Top             =   1170
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9022
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "MONTHLY"
      TabPicture(0)   =   "frm_rpt_layout_absen.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra_process_monthly"
      Tab(0).Control(1)=   "fra_monthly"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "PERIODE"
      TabPicture(1)   =   "frm_rpt_layout_absen.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fra_periode"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LynxGrid2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "RESERVED"
      TabPicture(2)   =   "frm_rpt_layout_absen.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame1 
         Caption         =   "List Karyawan"
         Height          =   3555
         Left            =   7110
         TabIndex        =   38
         Top             =   840
         Width           =   2895
         Begin VB.ListBox list_employee 
            Height          =   3180
            ItemData        =   "frm_rpt_layout_absen.frx":0060
            Left            =   90
            List            =   "frm_rpt_layout_absen.frx":0062
            TabIndex        =   39
            Top             =   240
            Width           =   2715
         End
      End
      Begin prj_tpc.LynxGrid LynxGrid2 
         Height          =   2055
         Left            =   1500
         TabIndex        =   37
         Top             =   2130
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3625
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
      Begin VB.Frame fra_monthly 
         Height          =   2655
         Left            =   -74190
         TabIndex        =   6
         Top             =   960
         Width           =   8415
         Begin VB.ComboBox cbo_monthly_company 
            Height          =   315
            ItemData        =   "frm_rpt_layout_absen.frx":0064
            Left            =   1800
            List            =   "frm_rpt_layout_absen.frx":006E
            TabIndex        =   12
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.Frame fra_monthly_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   8
            Top             =   960
            Width           =   4575
            Begin VB.CommandButton cmd_monthly_browse_employee 
               Caption         =   "..."
               Height          =   300
               Left            =   1440
               TabIndex        =   11
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txt_monthly_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   10
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox txt_monthly_employee_code 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   0
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   9
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.ComboBox cbo_monthly_employee 
            Height          =   315
            ItemData        =   "frm_rpt_layout_absen.frx":0081
            Left            =   1800
            List            =   "frm_rpt_layout_absen.frx":008B
            TabIndex        =   7
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_monthly 
            Height          =   300
            Left            =   1800
            TabIndex        =   13
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   178913283
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label13 
            Caption         =   "* yyyy-MM"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   3600
            TabIndex        =   32
            Top             =   1590
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Perusahaan"
            Height          =   195
            Left            =   720
            TabIndex        =   16
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Bulan"
            Height          =   195
            Left            =   720
            TabIndex        =   15
            Top             =   1620
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Karyawan"
            Height          =   195
            Left            =   720
            TabIndex        =   14
            Top             =   1200
            Width           =   705
         End
      End
      Begin VB.Frame fra_process_monthly 
         Height          =   2655
         Left            =   -73590
         TabIndex        =   21
         Top             =   930
         Width           =   8415
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Processing, Please Wait..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2640
            TabIndex        =   22
            Top             =   1080
            Width           =   2730
         End
      End
      Begin VB.Frame fra_periode 
         Height          =   3555
         Left            =   90
         TabIndex        =   17
         Top             =   840
         Width           =   6975
         Begin VB.TextBox txtProses 
            Height          =   315
            Left            =   1410
            TabIndex        =   47
            Top             =   3000
            Width           =   4155
         End
         Begin VB.TextBox txtSetuju 
            Height          =   315
            Left            =   1410
            TabIndex        =   45
            Top             =   2640
            Width           =   4155
         End
         Begin VB.TextBox txtTugas 
            Height          =   315
            Left            =   1410
            TabIndex        =   43
            Top             =   2280
            Width           =   4155
         End
         Begin VB.CommandButton cmdDelete 
            Appearance      =   0  'Flat
            Caption         =   "<<"
            Height          =   435
            Left            =   6150
            TabIndex        =   42
            Top             =   1350
            Width           =   645
         End
         Begin VB.CommandButton cmdInsert 
            Appearance      =   0  'Flat
            Caption         =   ">>"
            Height          =   435
            Left            =   6150
            TabIndex        =   41
            Top             =   900
            Width           =   645
         End
         Begin VB.TextBox txtkdkar 
            Height          =   285
            Left            =   3000
            TabIndex        =   40
            Text            =   "Text2"
            Top             =   180
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtnmkar 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            DragMode        =   1  'Automatic
            Height          =   285
            Left            =   2910
            TabIndex        =   36
            Top             =   990
            Width           =   2625
         End
         Begin VB.TextBox txt_nik 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1410
            TabIndex        =   35
            Top             =   990
            Width           =   1125
         End
         Begin VB.CommandButton vbbutton2 
            Caption         =   "..."
            Height          =   285
            Left            =   2550
            TabIndex        =   34
            Top             =   990
            Width           =   315
         End
         Begin VB.Frame fra_periode_department 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   3210
            TabIndex        =   28
            Top             =   450
            Width           =   3735
            Begin VB.TextBox txt_department_name 
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
               Height          =   315
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   50
               MultiLine       =   -1  'True
               TabIndex        =   29
               Top             =   90
               Width           =   2265
            End
            Begin TrueOleDBList60.TDBCombo TDBCombo_department 
               Height          =   375
               Left            =   0
               OleObjectBlob   =   "frm_rpt_layout_absen.frx":009C
               TabIndex        =   30
               Top             =   90
               Width           =   1305
            End
            Begin MSAdodcLib.Adodc Adodc_department 
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
         Begin VB.ComboBox cbo_periode_department 
            Height          =   315
            ItemData        =   "frm_rpt_layout_absen.frx":205D
            Left            =   1410
            List            =   "frm_rpt_layout_absen.frx":2067
            TabIndex        =   27
            Text            =   "..."
            Top             =   540
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   1410
            TabIndex        =   24
            Top             =   1410
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   177995779
            CurrentDate     =   39278
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Diproses Oleh"
            Height          =   195
            Left            =   150
            TabIndex        =   48
            Top             =   3030
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Disetujui Oleh"
            Height          =   195
            Left            =   150
            TabIndex        =   46
            Top             =   2670
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ditugaskan Oleh"
            Height          =   195
            Left            =   150
            TabIndex        =   44
            Top             =   2310
            Width           =   1185
         End
         Begin VB.Label Label14 
            Caption         =   "* yyyy-MM"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3180
            TabIndex        =   33
            Top             =   1440
            Width           =   1035
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Bulan"
            Height          =   195
            Left            =   150
            TabIndex        =   25
            Top             =   1410
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Dept / Area"
            Height          =   195
            Left            =   150
            TabIndex        =   19
            Top             =   570
            Width           =   840
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Karyawan"
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   1020
            Width           =   705
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   6450
      Width           =   10095
      Begin VB.CommandButton cmdPrint 
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
         Left            =   5250
         Picture         =   "frm_rpt_layout_absen.frx":207A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_print_sum_list 
         Caption         =   "&Sum. List"
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
         Left            =   5520
         Picture         =   "frm_rpt_layout_absen.frx":2604
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmd_send_mail 
         Caption         =   "Send &Mail"
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
         Left            =   6600
         Picture         =   "frm_rpt_layout_absen.frx":2B8E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   840
         Top             =   360
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
         Picture         =   "frm_rpt_layout_absen.frx":3118
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_rpt_layout_absen.frx":36A2
      TabIndex        =   4
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
      Caption         =   "LAPORAN PRINT OUT KEHADIRAN"
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
      Left            =   2745
      TabIndex        =   31
      Top             =   0
      Width           =   5355
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Perusahaan"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   780
      Width           =   855
   End
End
Attribute VB_Name = "frm_rpt_layout_absen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim strsql As String
Dim rs2 As New ADODB.Recordset

Private Sub cbo_monthly_company_Click()
If cbo_monthly_company.ListIndex = 0 Then
    cbo_monthly_employee.ListIndex = 0
    cbo_monthly_employee.Enabled = False
ElseIf cbo_monthly_company.ListIndex = 1 Then
    cbo_monthly_employee.ListIndex = 0
    cbo_monthly_employee.Enabled = True
End If
End Sub

Private Sub cbo_periode_department_Click()
TDBCombo_department.Text = ""
txt_department_name.Text = ""
    
If cbo_periode_department.ListIndex = 0 Then
    fra_periode_department.Visible = False
Else
    fra_periode_department.Visible = True
End If
End Sub


Private Sub cmd_print_slip_Click()
If check_validate_tdbcombo(TDBCombo_company) = False Then
    MsgBox "No Company selected!", vbInformation, headerMSG
    Exit Sub
End If

End Sub

Private Sub CmdExit_Click()
strsql = "DELETE from temp_list"
CnG.Execute strsql

Unload Me
End Sub

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

Private Sub cmdInsert_Click()
On Error Resume Next

CnG.BeginTrans
strsql = "INSERT into temp_list(employee_code,nik,employee_name) VALUES " _
        & "(" & txtkdkar.Text & ",'" & txt_nik & "','" & txtnmkar.Text & "')"
CnG.Execute strsql
CnG.CommitTrans

list_employee.Clear

strsql = "select employee_name from temp_list"
rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs2.RecordCount > 0 Then
    rs2.MoveFirst
    While Not rs2.EOF
        list_employee.AddItem rs2!EMPLOYEE_NAME
        rs2.MoveNext
    Wend
End If
rs2.Close

txtkdkar.Text = ""
txtnmkar.Text = ""
txt_nik.Text = ""
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next

Dim a As Integer

a = list_employee.ListIndex

CnG.BeginTrans
strsql = "DELETE from temp_list where employee_name = '" & list_employee.List(a) & "'"
CnG.Execute strsql
CnG.CommitTrans

list_employee.Clear

strsql = "select employee_name from temp_list"
rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs2.RecordCount > 0 Then
    rs2.MoveFirst
    While Not rs2.EOF
        list_employee.AddItem rs2!EMPLOYEE_NAME
        rs2.MoveNext
    Wend
End If
rs2.Close
End Sub

Private Sub cmdPrint_Click()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim year As Integer

str_file = "\report\rpt_layout_absen.rpt"

    str_sql = "SELECT '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'," _
            & "'" & TDBCombo_company.Text & "','" & txt_company_name.Text & "'," _
            & "'" & TDBCombo_department.Text & "','" & txt_department_name.Text & "'," _
            & "employee_code,nik,employee_name, '" & UCase(txtTugas.Text) & "'," _
            & "'" & UCase(txtSetuju.Text) & "','" & UCase(txtProses.Text) & "' " _
            & "FROM temp_list"

'str_param_periode = "MONTHLY : (" & Format(DTPicker_yearly.Value, "yyyy-MM") & ")"

Call a.Show
a.Caption = "PRINT OUT KEHADIRAN"""
Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub Form_Load()
Adodc_company.ConnectionString = strConn
Adodc_department.ConnectionString = strConn

Call load_data_company

Call createKar

DTPicker1.Value = Now

timer1.Enabled = True
SSTab1.TabVisible(0) = False

cbo_periode_department.ListIndex = 0

RemoveButtonX Me
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
End If

Call load_data_department
End Sub

Private Sub Timer1_Timer()
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
'If LOGIN_LEVEL = 100 Then
''    cbo_daily_company.Enabled = True
'    cbo_monthly_company.Enabled = True
'    cbo_periode_company.Enabled = True
'Else
''    cbo_daily_company.Enabled = False
'    cbo_monthly_company.Enabled = False
'    cbo_periode_company.Enabled = False
'End If
End Sub

Private Sub set_data_department(ByVal str_code As String)
On Error Resume Next

Adodc_department.Recordset.MoveFirst
Adodc_department.Recordset.Find ("department_code='" & str_code & "'")   ', 0, adSearchForward, 1)
If Not (Adodc_department.Recordset.EOF = True Or Adodc_department.Recordset.BOF = True) Then
    TDBCombo_department.Bookmark = Adodc_department.Recordset.AbsolutePosition
    Call TDBCombo_department_itemChange
Else
    TDBCombo_department.Text = ""
End If
End Sub

Private Sub load_data_department()
TDBCombo_department.Text = "": txt_department_name = ""

Adodc_department.RecordSource = "select * from m_department where company_code='" _
& TDBCombo_company.Columns("company_code").Value & "' order by department_code"
Adodc_department.Refresh

TDBCombo_department.RowSource = Adodc_department
End Sub

Private Sub TDBCombo_department_itemChange()
If TDBCombo_department.ApproxCount > 0 Then
    TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
    txt_department_name = TDBCombo_department.Columns("department_name").Value
End If
End Sub

Private Sub createKar()
   With LynxGrid2
      .AddColumn "NIK", 2000, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "employee_code", 2000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub
   
Private Sub isiGridKar(pilihan As Integer)
Dim v_department As String

If cbo_periode_department.ListIndex = 0 Then
    v_department = "AND "
Else
    v_department = "AND department_code = '" & TDBCombo_department.Text & "' AND "
End If

    If pilihan = 1 Then
        LynxGrid2.Clear
        strsql = "select nik,employee_name,employee_code " _
                & "from m_employee " _
                & "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " _
                & v_department _
                & "(employee_code LIKE '%" & txt_nik.Text & "%' " _
                & "OR employee_name LIKE '%" & txt_nik.Text & "%') " _
                & "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
                
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid2.AddItem rs2!nik & vbTab & rs2!EMPLOYEE_NAME & vbTab & rs2!employee_code
                rs2.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs2.RecordCount = 1 Then
                rs2.MoveFirst
                txtkdkar.Text = rs2!employee_code
                txtnmkar.Text = rs2!EMPLOYEE_NAME
                txt_nik.Text = rs2!nik
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
            txtkdkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub vbbutton2_Click()
    isiGridKar (1)
End Sub

Private Sub txt_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
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
