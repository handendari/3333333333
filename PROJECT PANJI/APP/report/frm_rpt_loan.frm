VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_rpt_loan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LAPORAN - PINJAMAN"
   ClientHeight    =   7440
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   10560
   Icon            =   "frm_rpt_loan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin prj_panji.LynxGrid LynxGrid2 
      Height          =   2925
      Left            =   4710
      TabIndex        =   21
      Top             =   3450
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
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Top             =   930
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   1530
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "BULANAN"
      TabPicture(0)   =   "frm_rpt_loan.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   870
         TabIndex        =   5
         Top             =   420
         Width           =   8415
         Begin VB.TextBox txt_division_name 
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
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   810
            Width           =   2655
         End
         Begin VB.Frame fra_periode_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   8
            Top             =   960
            Width           =   4575
            Begin VB.TextBox txt_periode_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               DragMode        =   1  'Automatic
               Height          =   285
               Left            =   1920
               TabIndex        =   19
               Top             =   240
               Width           =   2385
            End
            Begin VB.TextBox txt_periode_nik 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   0
               TabIndex        =   18
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox txt_periode_employee_code 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4170
               TabIndex        =   17
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
            Begin prj_panji.vbButton cmd_periode_browse_employee 
               Height          =   285
               Left            =   1560
               TabIndex        =   20
               Top             =   240
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
               MICON           =   "frm_rpt_loan.frx":05A6
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
            ItemData        =   "frm_rpt_loan.frx":05C2
            Left            =   1800
            List            =   "frm_rpt_loan.frx":05CC
            TabIndex        =   7
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "DAY COUNT"
            Height          =   495
            Left            =   0
            TabIndex        =   6
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker DTPicker_monthly 
            Height          =   300
            Left            =   1800
            TabIndex        =   9
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   107610115
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_division 
            Height          =   375
            Left            =   1800
            OleObjectBlob   =   "frm_rpt_loan.frx":05DD
            TabIndex        =   16
            Top             =   840
            Width           =   1695
         End
         Begin MSAdodcLib.Adodc Adodc_division 
            Height          =   375
            Left            =   2490
            Top             =   810
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
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KARYAWAN"
            Height          =   195
            Left            =   765
            TabIndex        =   12
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DIVISI"
            Height          =   195
            Left            =   1215
            TabIndex        =   11
            Top             =   870
            Width           =   465
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BULAN"
            Height          =   195
            Left            =   1155
            TabIndex        =   10
            Top             =   1590
            Width           =   540
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   5970
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   840
         Top             =   360
      End
      Begin prj_panji.vbButton cmdPrint 
         Height          =   705
         Left            =   4800
         TabIndex        =   22
         Top             =   300
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
         MICON           =   "frm_rpt_loan.frx":259C
         PICN            =   "frm_rpt_loan.frx":25B8
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
         Left            =   5790
         TabIndex        =   23
         Top             =   300
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
         MICON           =   "frm_rpt_loan.frx":364A
         PICN            =   "frm_rpt_loan.frx":3666
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdExit 
         Height          =   705
         Left            =   8400
         TabIndex        =   24
         Top             =   330
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
         MICON           =   "frm_rpt_loan.frx":46F8
         PICN            =   "frm_rpt_loan.frx":4714
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
      OleObjectBlob   =   "frm_rpt_loan.frx":57A6
      TabIndex        =   3
      Top             =   930
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1080
      Top             =   930
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
      Top             =   630
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "LAPORAN PINJAMAN"
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
      Left            =   360
      TabIndex        =   14
      Top             =   210
      Width           =   4365
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   930
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_rpt_loan.frx":7764
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14850
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CATATAN AKTIFITAS"
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
      Left            =   300
      TabIndex        =   13
      Top             =   150
      Width           =   4365
   End
End
Attribute VB_Name = "frm_rpt_loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim int_libur() As Integer

Private Sub cbo_periode_employee_Click()
    If cbo_periode_employee.ListIndex = 0 Then
        fra_periode_employee.Visible = False
    Else
        fra_periode_employee.Visible = True
    End If
    
    txt_periode_employee_code = "": txt_periode_employee_name = "": txt_periode_nik = ""
End Sub

Private Sub cmdSummary_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        If check_valid_periode Then
            Call rpt_monthly_summary
        End If
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Function check_valid_periode() As Boolean
check_valid_periode = True

    'validate employee
    If cbo_periode_employee.ListIndex = 1 And Trim(txt_periode_employee_code) = "" Then
        MsgBox "Karyawan Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
        cmd_periode_browse_employee.SetFocus
        check_valid_periode = False
        Exit Function
    End If
End Function

Private Sub rpt_monthly_summary()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim year As Integer

    str_file = "\report\rpt_loan_summary.rpt"
    
    If cbo_periode_employee.ListIndex = 0 Then
        If LOGIN_LEVEL = 100 Then
            str_sql = "SELECT c.company_name,e.division_name,f.title_name, " _
                        & "b.nik,b.employee_name,a.installment_month,a.sequence_number, " _
                        & "a.installment_amount , a.flag_paid, " _
                        & "IFNULL((SELECT SUM(a.installment_amount) " _
                            & "FROM td_loan WHERE employee_code = b.employee_code " _
                            & "AND left(installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' " _
                            & "AND flag_paid = 1),0) paid, " _
                        & "COUNT(DISTINCT a.installment_amount) count_loan, " _
                        & "(SUM(DISTINCT a.installment_amount)*COUNT(DISTINCT a.installment_amount)) loan_total " _
                    & "FROM td_loan a JOIN m_employee b ON a.employee_code = b.employee_code " _
                    & "JOIN m_company c ON b.company_code = c.company_code " _
                    & "JOIN tm_loan d ON a.employee_code = d.employee_code " _
                    & "JOIN m_division e ON b.division_code = e.division_code and b.company_code = e.company_code " _
                    & "JOIN m_title f ON b.title_code = f.title_code " _
                    & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                                "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                    & "AND left(a.installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' " _
                    & "GROUP BY a.employee_code"
        Else
            str_sql = "SELECT c.company_name,e.division_name,f.title_name, " _
                        & "b.nik,b.employee_name,a.installment_month,a.sequence_number, " _
                        & "a.installment_amount , a.flag_paid, " _
                        & "IFNULL((SELECT SUM(a.installment_amount) " _
                            & "FROM td_loan WHERE employee_code = b.employee_code " _
                            & "AND left(installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' " _
                            & "AND flag_paid = 1),0) paid, " _
                        & "COUNT(DISTINCT a.installment_amount) count_loan, " _
                        & "(SUM(DISTINCT a.installment_amount)*COUNT(DISTINCT a.installment_amount)) loan_total " _
                    & "FROM td_loan a JOIN m_employee b ON a.employee_code = b.employee_code " _
                    & "JOIN m_company c ON b.company_code = c.company_code " _
                    & "JOIN tm_loan d ON a.employee_code = d.employee_code " _
                    & "JOIN m_division e ON b.division_code = e.division_code and b.company_code = e.company_code " _
                    & "JOIN m_title f ON b.title_code = f.title_code " _
                    & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                                "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                    & "AND left(a.installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' AND " _
                    & "(level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                    & "GROUP BY a.employee_code"
        End If
    Else
        str_sql = "SELECT c.company_name,e.division_name,f.title_name, " _
                    & "b.nik,b.employee_name,a.installment_month,a.sequence_number, " _
                    & "a.installment_amount , a.flag_paid, " _
                    & "IFNULL((SELECT SUM(a.installment_amount) " _
                            & "FROM td_loan WHERE employee_code = b.employee_code " _
                            & "AND left(installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' " _
                            & "AND flag_paid = 1),0) paid, " _
                    & "COUNT(DISTINCT a.installment_amount) count_loan, " _
                    & "(SUM(DISTINCT a.installment_amount)*COUNT(DISTINCT a.installment_amount)) loan_total " _
                & "FROM td_loan a JOIN m_employee b ON a.employee_code = b.employee_code " _
                & "JOIN m_company c ON b.company_code = c.company_code " _
                & "JOIN tm_loan d ON a.employee_code = d.employee_code " _
                & "JOIN m_division e ON b.division_code = e.division_code and b.company_code = e.company_code " _
                & "JOIN m_title f ON b.title_code = f.title_code " _
                & "WHERE b.employee_code = '" & txt_periode_employee_code & "' AND b.company_code = '" & TDBCombo_company.Text & "' " _
                & "AND left(a.installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' " _
                & "GROUP BY a.employee_code"
    End If
    
    str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
    
    Call a.Show
    a.Caption = "REPORT EMPLOYEE LOAN"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub rpt_monthly()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim year As Integer

    str_file = "\report\rpt_loan.rpt"
    
    If cbo_periode_employee.ListIndex = 0 Then
        If LOGIN_LEVEL = 100 Then
            str_sql = "SELECT c.company_name,e.division_name,f.title_name, " _
                        & "b.nik,b.employee_name,a.installment_month,a.sequence_number, " _
                        & "a.installment_amount , a.flag_paid, " _
                        & "IFNULL((SELECT SUM(a.installment_amount) " _
                            & "FROM td_loan WHERE employee_code = b.employee_code " _
                            & "AND left(installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' " _
                            & "AND flag_paid = 1),0) paid, " _
                        & "COUNT(DISTINCT a.installment_amount) count_loan, " _
                        & "(SUM(DISTINCT a.installment_amount)*COUNT(DISTINCT a.installment_amount)) loan_total " _
                    & "FROM td_loan a JOIN m_employee b ON a.employee_code = b.employee_code " _
                    & "JOIN m_company c ON b.company_code = c.company_code " _
                    & "JOIN tm_loan d ON a.employee_code = d.employee_code " _
                    & "JOIN m_division e ON b.division_code = e.division_code and b.company_code = e.company_code " _
                    & "JOIN m_title f ON b.title_code = f.title_code " _
                    & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                            "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                    & "AND left(a.installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' " _
                    & "GROUP BY a.employee_code"
        Else
            str_sql = "SELECT c.company_name,e.division_name,f.title_name, " _
                        & "b.nik,b.employee_name,a.installment_month,a.sequence_number, " _
                        & "a.installment_amount , a.flag_paid, " _
                        & "IFNULL((SELECT SUM(a.installment_amount) " _
                            & "FROM td_loan WHERE employee_code = b.employee_code " _
                            & "AND left(installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' " _
                            & "AND flag_paid = 1),0) paid, " _
                        & "COUNT(DISTINCT a.installment_amount) count_loan, " _
                        & "(SUM(DISTINCT a.installment_amount)*COUNT(DISTINCT a.installment_amount)) loan_total " _
                    & "FROM td_loan a JOIN m_employee b ON a.employee_code = b.employee_code " _
                    & "JOIN m_company c ON b.company_code = c.company_code " _
                    & "JOIN tm_loan d ON a.employee_code = d.employee_code " _
                    & "JOIN m_division e ON b.division_code = e.division_code and b.company_code = e.company_code " _
                    & "JOIN m_title f ON b.title_code = f.title_code " _
                    & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                            "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                    & "AND left(a.installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' AND " _
                    & "(level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                    & "GROUP BY a.employee_code"

        End If
    Else
        str_sql = "SELECT c.company_name,e.division_name,f.title_name, " _
                    & "b.nik,b.employee_name,a.installment_month,a.sequence_number, " _
                    & "a.installment_amount , a.flag_paid, " _
                    & "IFNULL((SELECT SUM(a.installment_amount) " _
                            & "FROM td_loan WHERE employee_code = b.employee_code " _
                            & "AND left(installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' " _
                            & "AND flag_paid = 1),0) paid, " _
                    & "COUNT(DISTINCT a.installment_amount) count_loan, " _
                    & "(SUM(DISTINCT a.installment_amount)*COUNT(DISTINCT a.installment_amount)) loan_total " _
                & "FROM td_loan a JOIN m_employee b ON a.employee_code = b.employee_code " _
                & "JOIN m_company c ON b.company_code = c.company_code " _
                & "JOIN tm_loan d ON a.employee_code = d.employee_code " _
                & "JOIN m_division e ON b.division_code = e.division_code and b.company_code = e.company_code " _
                & "JOIN m_title f ON b.title_code = f.title_code " _
                & "WHERE b.employee_code = '" & txt_periode_employee_code & "' AND b.company_code = '" & TDBCombo_company.Text & "' " _
                & "AND left(a.installment_month,7) = '" & Format(DTPicker_monthly.Value, "yyyy-MM") & "' " _
                & "GROUP BY a.employee_code"
    End If
    
    str_param_periode = "MONTHLY : (" & Format(DTPicker_monthly.Value, "yyyy-MM") & ")"
    
    Call a.Show
    a.Caption = "REPORT EMPLOYEE LOAN"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub cmdPrint_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        If check_valid_periode Then
            Call rpt_monthly
        End If
    'ElseIf SSTab1.Tab = 0 Then
    '    If check_valid_yearly Then
    '        Call rpt_yearly
    '    End If
    End If
End Sub

Private Sub Form_Load()
    Adodc_company.ConnectionString = strConn
    Adodc_division.ConnectionString = strConn
    
    Call load_data_company
    Call load_data_user_access(Me)
    Call createGridKar
    
    DTPicker_monthly.Value = Now
    
    cbo_periode_employee.ListIndex = 0
    
    Timer1.Enabled = True
    SSTab1.Tab = 0
End Sub

Private Sub load_data_company()
    Adodc_company.RecordSource = "select * from m_company order by company_code"
    Adodc_company.Refresh
    
    TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub load_data_division()
    Adodc_division.RecordSource = "select * from m_division where company_code = '" & TDBCombo_company.Text & "' order by division_code"
    Adodc_division.Refresh
    
    TDBCombo_division.RowSource = Adodc_division
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
    End If
    
    Call load_data_division
End Sub

Private Sub TDBCombo_division_itemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
        txt_division_name.Text = TDBCombo_division.Columns("division_name").Value
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call set_company_mode_adodc(Adodc_company, TDBCombo_company, txt_company_name)
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
                    & "WHERE " & IIf(TDBCombo_division.Text = "", "a.company_code = '" & TDBCombo_company.Text & "'", _
                            "a.company_code = '" & TDBCombo_company.Text & "' AND a.division_code = '" & TDBCombo_division.Text & "'") & " " _
                        & "AND (a.nik LIKE '%" & txt_periode_nik.Text & "%' " _
                        & "OR a.employee_name LIKE '%" & txt_periode_nik.Text & "%') " _
                        & "AND a.flag_active <> 0"
        Else
            SQL = "SELECT a.nik,a.employee_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE " & IIf(TDBCombo_division.Text = "", "a.company_code = '" & TDBCombo_company.Text & "'", _
                            "a.company_code = '" & TDBCombo_company.Text & "' AND a.division_code = '" & TDBCombo_division.Text & "'") & " " _
                        & "AND (a.nik LIKE '%" & txt_periode_nik.Text & "%' " _
                        & "OR a.employee_name LIKE '%" & txt_periode_nik.Text & "%') " _
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
                txt_periode_employee_code.Text = rs!employee_code
                txt_periode_employee_name.Text = rs!EMPLOYEE_NAME
                txt_periode_nik.Text = rs!nik
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
            txt_periode_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txt_periode_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txt_periode_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
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

Private Sub txt_periode_nik_Change()
    If txt_periode_nik.Text = "" Then
        txt_periode_employee_code.Text = ""
        txt_periode_employee_name.Text = ""
    End If
End Sub

Private Sub txt_periode_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmd_periode_browse_employee_Click()
    isiGridKar (1)
End Sub

