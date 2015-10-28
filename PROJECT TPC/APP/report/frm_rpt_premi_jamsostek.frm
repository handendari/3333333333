VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_rpt_premi_jamsostek 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORT PREMI JAMSOSTEK"
   ClientHeight    =   7305
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   10560
   Icon            =   "frm_rpt_premi_jamsostek.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   3465
      Left            =   4710
      TabIndex        =   29
      Top             =   3510
      Visible         =   0   'False
      Width           =   5025
      _ExtentX        =   8864
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
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3300
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   870
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   240
      TabIndex        =   2
      Top             =   1350
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "PERIODE"
      TabPicture(0)   =   "frm_rpt_premi_jamsostek.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   870
         TabIndex        =   5
         Top             =   630
         Width           =   8415
         Begin VB.ComboBox cbo_periode_to 
            Height          =   315
            ItemData        =   "frm_rpt_premi_jamsostek.frx":05A6
            Left            =   3600
            List            =   "frm_rpt_premi_jamsostek.frx":05B0
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "..."
            Top             =   1950
            Width           =   1335
         End
         Begin VB.ComboBox cbo_periode_department 
            Height          =   315
            ItemData        =   "frm_rpt_premi_jamsostek.frx":05BE
            Left            =   1800
            List            =   "frm_rpt_premi_jamsostek.frx":05C8
            TabIndex        =   17
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.Frame fra_periode_department 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   3600
            TabIndex        =   14
            Top             =   750
            Width           =   4695
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
               Height          =   285
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               MultiLine       =   -1  'True
               TabIndex        =   15
               Top             =   90
               Width           =   2415
            End
            Begin TrueOleDBList60.TDBCombo TDBCombo_department 
               Height          =   375
               Left            =   0
               OleObjectBlob   =   "frm_rpt_premi_jamsostek.frx":05DB
               TabIndex        =   16
               Top             =   90
               Width           =   1815
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
         Begin VB.ComboBox cbo_periode_company 
            Height          =   315
            ItemData        =   "frm_rpt_premi_jamsostek.frx":2544
            Left            =   3600
            List            =   "frm_rpt_premi_jamsostek.frx":254E
            TabIndex        =   13
            Text            =   "..."
            Top             =   1560
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Frame fra_periode_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   9
            Top             =   960
            Width           =   4575
            Begin VB.TextBox txt_periode_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   11
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox txt_periode_nik 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               Height          =   315
               Left            =   0
               MaxLength       =   50
               TabIndex        =   10
               Top             =   240
               Width           =   1335
            End
            Begin prj_tpc.vbButton cmd_periode_browse_employee 
               Height          =   315
               Left            =   1410
               TabIndex        =   28
               Top             =   240
               Width           =   405
               _ExtentX        =   714
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
               MICON           =   "frm_rpt_premi_jamsostek.frx":2561
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
            ItemData        =   "frm_rpt_premi_jamsostek.frx":257D
            Left            =   1800
            List            =   "frm_rpt_premi_jamsostek.frx":2587
            TabIndex        =   8
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "DAY COUNT"
            Height          =   495
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txt_periode_employee_code 
            Height          =   285
            Left            =   4290
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   390
            Visible         =   0   'False
            Width           =   315
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_from 
            Height          =   300
            Left            =   1800
            TabIndex        =   20
            Top             =   1950
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   96141315
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_to 
            Height          =   300
            Left            =   5040
            TabIndex        =   21
            Top             =   1950
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   96141315
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   1800
            TabIndex        =   22
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "MM-yyyy"
            Format          =   96141315
            CurrentDate     =   39278
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PERIODE"
            Height          =   195
            Left            =   960
            TabIndex        =   24
            Top             =   1980
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "MONTH"
            Height          =   195
            Left            =   1080
            TabIndex        =   23
            Top             =   1560
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "DEPARTMENT"
            Height          =   195
            Left            =   570
            TabIndex        =   18
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   810
            TabIndex        =   12
            Top             =   1200
            Width           =   870
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   5790
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   840
         Top             =   360
      End
      Begin prj_tpc.vbButton cmdExit 
         Height          =   705
         Left            =   8640
         TabIndex        =   25
         Top             =   300
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
         MICON           =   "frm_rpt_premi_jamsostek.frx":2598
         PICN            =   "frm_rpt_premi_jamsostek.frx":25B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdPrint 
         Height          =   705
         Left            =   6480
         TabIndex        =   26
         Top             =   300
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Slip"
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
         MICON           =   "frm_rpt_premi_jamsostek.frx":3646
         PICN            =   "frm_rpt_premi_jamsostek.frx":3662
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
      Left            =   1560
      OleObjectBlob   =   "frm_rpt_premi_jamsostek.frx":46F4
      TabIndex        =   1
      Top             =   870
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1440
      Top             =   840
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
      Left            =   8970
      Top             =   870
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT PREMI JAMSOSTEK"
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
      TabIndex        =   27
      Top             =   180
      Width           =   4155
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   930
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_rpt_premi_jamsostek.frx":665A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12690
   End
End
Attribute VB_Name = "frm_rpt_premi_jamsostek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim int_libur() As Integer

Private Sub cbo_periode_department_Click()
    TDBCombo_department.Text = ""
    txt_department_name.Text = ""
        
    If cbo_periode_department.ListIndex = 0 Then
        fra_periode_department.Visible = False
    Else
        fra_periode_department.Visible = True
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

'Private Sub cmd_periode_browse_employee_Click()
'    frm_lookup_mst_employee.public_int_mode = 168
'    frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
'    frm_lookup_mst_employee.Show 1
'End Sub

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

Private Sub cmdPrint_Click()
Dim tgl1 As String, tgl2 As String, kdkaryawan As String
Dim a As New frm_rpt

    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
    int_process = vbNo
    
    str_file = "\report\rpt_premi_jamsostek.rpt"
    
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
    int_flag_employee = 0
    
    d1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
    d2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
    
    If cbo_periode_employee.Text = "All" Then
        str_employee_code = ""
    Else
        str_employee_code = txt_periode_employee_code.Text
    End If
    
    strsql = "call sp_premi_jamsostek('" & d1 & "','" & d2 & "'," _
            & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_department.Text & "'," _
            & int_flag_employee & ",'" & str_employee_code & "'," _
            & cbo_periode_department.ListIndex & ",'" & LOGIN_LEVEL & "')"
    
    Call a.Show
    a.Caption = "SUMMARY PREMI JAMSOSTEK"
    'str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
     
    Call a.rpt_view(strsql, str_file, pTgl1)

End Sub

Private Sub Form_Load()
    Adodc_company.ConnectionString = strConn
    Adodc_department.ConnectionString = strConn
    
    Call load_data_company
    Call load_data_user_access(Me)
    
    DTPicker_periode_from.Value = Now
    DTPicker_periode_to.Value = Now
    DTPicker1.Value = Now
    
    Call getPeriodeProses
    
    cbo_periode_to.ListIndex = 0
    cbo_periode_company.ListIndex = 1
    cbo_periode_employee.ListIndex = 0
    
'    cbo_monthly_company.ListIndex = 1
    'cbo_monthly_employee.ListIndex = 0
    
    timer1.Enabled = True
    cbo_periode_department.ListIndex = 0
    
    Call createGridKar
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
    Call set_company_mode_adodc(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Private Sub set_data_department(ByVal str_code As String)
On Error Resume Next

    Adodc_department.Recordset.MoveFirst
    Adodc_department.Recordset.Find ("department_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (Adodc_department.Recordset.EOF = True Or Adodc_department.Recordset.BOF = True) Then
        TDBCombo_department.Bookmark = Adodc_department.Recordset.AbsolutePosition
        Call tdbCombo_department_itemChange
    Else
        TDBCombo_department.Text = ""
    End If
End Sub

Private Sub load_data_department()
TDBCombo_department.Text = "": txt_department_name = ""

    Adodc_department.RecordSource = "select department_code, department_name from m_department where company_code='" _
    & TDBCombo_company.Columns("company_code").Value & "' order by department_code"
    Adodc_department.Refresh
    
    TDBCombo_department.RowSource = Adodc_department
End Sub

Private Sub tdbCombo_department_itemChange()
    If TDBCombo_department.ApproxCount > 0 Then
        TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
        txt_department_name = TDBCombo_department.Columns("department_name").Value
    End If
End Sub

Private Sub DTPicker1_Change()
    Call getPeriodeProses
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    Call getPeriodeProses
End Sub

Private Sub getPeriodeProses()
Dim strsql As String
Dim rsbulan As New ADODB.Recordset

    strsql = "select a.date_from,a.date_to from h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " _
            & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " _
            & "AND left(`month`,7) = '" & Format(DTPicker1.Value, "yyyy-MM") & "'"
    rsbulan.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsbulan.RecordCount > 0 Then
        DTPicker_periode_from.Value = rsbulan!date_from
        DTPicker_periode_to.Value = rsbulan!date_to
        cbo_periode_to.ListIndex = 1
    End If
End Sub

    
Private Sub createGridKar()
   With LynxGrid2
      .AddColumn "Employee Code", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "employee_code", 2000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
Dim vParam As String
    If pilihan = 1 Then
        LynxGrid2.Clear
        
        vParam = IIf(DEPARTMENT_CODE <> "" And DIVISION_CODE = "", "a.department_code = '" & DEPARTMENT_CODE & "'", IIf(DEPARTMENT_CODE = "" And DIVISION_CODE = "", "a.company_code = '" & COMPANY_CODE & "'", "a.department_code = '" & DEPARTMENT_CODE & "' AND a.division_code = '" & DIVISION_CODE & "'"))
        
        If LOGIN_LEVEL = 100 Then
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND (nik LIKE '%" & txt_periode_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_periode_nik.Text & "%')"
        Else
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND " & vParam & " " & _
                        "AND (nik LIKE '%" & txt_periode_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_periode_nik.Text & "%') " & _
                        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
        End If
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid2.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME & vbTab & rs!employee_code
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
            txt_periode_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
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


