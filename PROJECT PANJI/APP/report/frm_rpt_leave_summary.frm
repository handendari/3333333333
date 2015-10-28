VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_rpt_summary_leave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SUMMARY LEAVE REPORT"
   ClientHeight    =   7290
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   10560
   Icon            =   "frm_rpt_leave_summary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3210
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Top             =   900
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   1380
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "PARAMETER"
      TabPicture(0)   =   "frm_rpt_leave_summary.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   2655
         Left            =   840
         TabIndex        =   5
         Top             =   900
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
            Left            =   3360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   930
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker DTPicker_summary_leave 
            Height          =   300
            Left            =   3360
            TabIndex        =   6
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   209584131
            CurrentDate     =   39278
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_division 
            Height          =   375
            Left            =   3360
            OleObjectBlob   =   "frm_rpt_leave_summary.frx":05A6
            TabIndex        =   12
            Top             =   570
            Width           =   1695
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DIVISI"
            Height          =   195
            Left            =   2580
            TabIndex        =   13
            Top             =   600
            Width           =   465
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TANGGAL"
            Height          =   195
            Left            =   2280
            TabIndex        =   7
            Top             =   1350
            Width           =   765
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   5820
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   840
         Top             =   360
      End
      Begin prj_panji.vbButton cmdExit 
         Height          =   705
         Left            =   8370
         TabIndex        =   9
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
         MICON           =   "frm_rpt_leave_summary.frx":2565
         PICN            =   "frm_rpt_leave_summary.frx":2581
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
         Left            =   7350
         TabIndex        =   10
         Top             =   300
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Cetak"
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
         MICON           =   "frm_rpt_leave_summary.frx":3613
         PICN            =   "frm_rpt_leave_summary.frx":362F
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
      Left            =   1410
      OleObjectBlob   =   "frm_rpt_leave_summary.frx":46C1
      TabIndex        =   3
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "LAPORAN REKAP CUTI"
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
      TabIndex        =   8
      Top             =   150
      Width           =   4365
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   930
      Width           =   1110
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_rpt_leave_summary.frx":667F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14850
   End
End
Attribute VB_Name = "frm_rpt_summary_leave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCompany As New ADODB.Recordset
Dim rsDiv As New ADODB.Recordset

Dim int_libur() As Integer
Dim v_access As String

Private Sub rpt_summary_leave()
    Dim str_param_periode, str_file As String
    Dim a As New frm_rpt
    
    str_file = "\report\rpt_summary_leave.rpt"
    v_access = IIf(LOGIN_LEVEL = 100, "", "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND flag_active = 0 order by employee_name")
    
    SQL = "select start_periode,end_periode,CAST(CONCAT(YEAR(b.start_periode),_latin1'-',YEAR(b.end_periode)) AS CHAR) AS leave_periode," & _
                "max_leave,actual_leave,over_leave,(-(1) * b.over_leave) AS leave_available,flag_close,nik,b.employee_code," & _
                "employee_name,employee_nick_name,a.division_code,e.division_name," & _
                "a.company_code,c.company_name,a.flag_active " & _
              "from m_employee a LEFT JOIN t_leave_periode b ON a.employee_code = b.employee_code " & _
                "JOIN m_company c ON a.company_code = c.company_code " & _
                "JOIN m_division e ON a.division_code = e.division_code " & _
                    "AND a.company_code = e.company_code " & _
                "JOIN m_title f ON a.title_code = f.title_code " & _
              "where " & IIf(TDBCombo_division.Text <> "", "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND a.division_code = '" & TDBCombo_division.Text & "'", "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "'") & " " & _
                "AND DATE(start_periode) <= '" & Format(DTPicker_summary_leave, "yyyy-MM-dd") & "' " & _
                "OR DATE(end_periode) >= '" & Format(DTPicker_summary_leave, "yyyy-MM-dd") & "' " & _
                "" & access & ""
    
    str_param_periode = "DATE : ( " & Format(DTPicker_summary_leave, "yyyy-MM-dd") & " )"
    
    Call a.Show
    a.Caption = "SUMMARY LEAVE REPORT"
    Call a.rpt_view(SQL, str_file, str_param_periode)
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
    Call rpt_summary_leave
    
End Sub

Private Sub Form_Load()
    Call load_data_company
    Call load_data_user_access(Me)
    
    DTPicker_summary_leave.Value = Now
    
    Timer1.Enabled = True
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If
    
    Call load_data_division
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub load_data_division()
    If rsDiv.State Then rsDiv.Close
    SQL = "select * from m_division where company_code = '" & TDBCombo_company.Text & "' order by division_code"
    rsDiv.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_division.RowSource = rsDiv
End Sub

Private Sub tdbcombo_division_Change()
    If TDBCombo_division.Text = "" Then txt_division_name.Text = ""
End Sub

Private Sub tdbcombo_division_itemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
        txt_division_name = TDBCombo_division.Columns("division_name").Value
    End If
End Sub
