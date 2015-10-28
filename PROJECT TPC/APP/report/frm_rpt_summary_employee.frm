VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Begin VB.Form frm_rpt_summary_employee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SUMMARY EMPLOYEE"
   ClientHeight    =   7095
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   10560
   Icon            =   "frm_rpt_summary_employee.frx":0000
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
      Left            =   3030
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Top             =   750
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   1230
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "SUMMARY EMPLOYEE"
      TabPicture(0)   =   "frm_rpt_summary_employee.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   2775
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   8415
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
            Left            =   3720
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   930
            Width           =   2655
         End
         Begin VB.TextBox txt_wt_to 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4770
            TabIndex        =   15
            Top             =   2010
            Width           =   615
         End
         Begin VB.TextBox txt_wt_from 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3720
            TabIndex        =   14
            Top             =   2010
            Width           =   615
         End
         Begin VB.CheckBox chk_wt 
            Caption         =   "WORKING AGE"
            Height          =   255
            Left            =   1800
            TabIndex        =   13
            Top             =   2010
            Width           =   1695
         End
         Begin VB.TextBox txt_age_to 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4770
            TabIndex        =   11
            Top             =   1650
            Width           =   615
         End
         Begin VB.TextBox txt_age_from 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3720
            TabIndex        =   10
            Top             =   1650
            Width           =   615
         End
         Begin VB.CheckBox chk_age 
            Caption         =   "AGE"
            Height          =   255
            Left            =   1800
            TabIndex        =   9
            Top             =   1650
            Width           =   1215
         End
         Begin VB.ComboBox cbo_sex 
            Height          =   315
            ItemData        =   "frm_rpt_summary_employee.frx":05A6
            Left            =   3720
            List            =   "frm_rpt_summary_employee.frx":05B0
            TabIndex        =   8
            Text            =   "..."
            Top             =   1290
            Width           =   1695
         End
         Begin VB.CheckBox chk_sex 
            Caption         =   "SEX"
            Height          =   255
            Left            =   1800
            TabIndex        =   7
            Top             =   1290
            Width           =   1695
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_department 
            Height          =   375
            Left            =   3720
            OleObjectBlob   =   "frm_rpt_summary_employee.frx":05C2
            TabIndex        =   18
            Top             =   570
            Width           =   1695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4410
            TabIndex        =   16
            Top             =   2040
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4410
            TabIndex        =   12
            Top             =   1680
            Width           =   270
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "DEPARTMENT"
            Height          =   195
            Left            =   1800
            TabIndex        =   6
            Top             =   600
            Width           =   1125
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   5670
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   30
         Top             =   150
      End
      Begin prj_tpc.vbButton cmdExit 
         Height          =   705
         Left            =   8340
         TabIndex        =   20
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
         MICON           =   "frm_rpt_summary_employee.frx":252B
         PICN            =   "frm_rpt_summary_employee.frx":2547
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
         Left            =   7320
         TabIndex        =   21
         Top             =   300
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Print"
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
         MICON           =   "frm_rpt_summary_employee.frx":35D9
         PICN            =   "frm_rpt_summary_employee.frx":35F5
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
      Left            =   1230
      OleObjectBlob   =   "frm_rpt_summary_employee.frx":4687
      TabIndex        =   3
      Top             =   750
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT SUMMARY EMPLOYEE"
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
      TabIndex        =   19
      Top             =   150
      Width           =   4365
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY :"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   810
      Width           =   885
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_rpt_summary_employee.frx":65ED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14850
   End
End
Attribute VB_Name = "frm_rpt_summary_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCompany As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset

Dim v_access As String
Dim v_sex As String
Dim v_age As String
Dim v_working_age As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub rpt_summary_employee()
Dim str_file As String
Dim a As New frm_rpt

    str_file = "\report\rpt_master_employee.rpt"
    
    v_access = IIf(LOGIN_LEVEL = 100, "", "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND flag_active = 0 order by employee_name")
    
    v_sex = IIf(chk_sex.Value = 0, "", IIf(chk_sex.Value And cbo_sex.ListIndex = 0, "WHERE sex = 0", " WHERE sex = 1"))
    v_age = IIf(chk_age.Value = 0, "", IIf(chk_age.Value And txt_age_to.Text <> "", "AND umur BETWEEN '" & txt_age_from.Text & "' AND '" & txt_age_to.Text & "'", "AND umur = '" & txt_age_from & "'"))
    v_working_age = IIf(chk_wt.Value = 0, "", IIf(chk_wt.Value And txt_wt_to.Text <> "", "AND lama_bekerja BETWEEN '" & txt_wt_from.Text & "' AND '" & txt_wt_to.Text & "'", "AND lama_bekerja = '" & txt_wt_from & "'"))
    
    SQL = "SELECT * FROM (SELECT a.company_code,b.company_name,a.department_code,c.department_name," & _
                   "a.employee_code,a.employee_name,a.title_code,e.title_name,a.marital_status," & _
                   "a.sex,a.start_working,a.bank_account,a.flag_active,a.division_code," & _
                   "d.division_name,a.no_of_children,a.religion,a.npwp,a.phone_number," & _
                   "a.place_birth,a.date_birth,a.emp_address, ROUND(DATEDIFF(NOW(),date_birth)/365) umur," & _
                   "ROUND(DATEDIFF(NOW(),start_working)/365) lama_bekerja " & _
                "FROM m_employee a LEFT JOIN m_company b ON a.company_code = b.company_code " & _
                "LEFT JOIN m_department c ON a.department_code = c.department_code AND c.company_code = b.company_code " & _
                "LEFT JOIN m_division d ON a.division_code = d.division_code AND d.department_code = c.department_code AND d.company_code = b.company_code " & _
                "LEFT JOIN m_title e ON a.title_code = e.title_code " & _
                "WHERE " & IIf(TDBCombo_department.Text <> "", "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND a.department_code = '" & TDBCombo_department.Text & "'", "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "'") & " " & _
                    "" & access & " " & _
                "ORDER BY employee_name ASC) aa " & _
          "" & v_sex & " " & v_age & " " & v_working_age & ""
                
    Call a.Show
    
    a.Caption = "REPORT SUMMARY EMPLOYEE"
    Call a.rpt_view(SQL, str_file, Now)
End Sub

Private Sub cmdPrint_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "No Company selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        Call rpt_summary_employee
    End If
End Sub

Private Sub Form_Load()

    Call load_data_company
    Call load_data_user_access(Me)
    
    timer1.Enabled = True
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub load_data_department()
    If rsDept.State Then rsDept.Close
    SQL = "select * from m_department where company_code = '" & TDBCombo_company.Text & "' order by department_code"
    rsDept.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_department.RowSource = rsDept
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If
    
    Call load_data_department
End Sub

Private Sub TDBCombo_department_Change()
    If TDBCombo_department.Text = "" Then txt_department_name.Text = ""
End Sub

Private Sub tdbCombo_department_itemChange()
    If TDBCombo_department.ApproxCount > 0 Then
        TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
        txt_department_name = TDBCombo_department.Columns("department_name").Value
    End If
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub
