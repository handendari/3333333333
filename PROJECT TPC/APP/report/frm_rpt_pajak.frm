VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_rpt_pajak 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORT SSP"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10410
   Icon            =   "frm_rpt_pajak.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   2940
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   0
      Top             =   900
      Width           =   3975
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1140
      OleObjectBlob   =   "frm_rpt_pajak.frx":058A
      TabIndex        =   6
      Top             =   900
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   960
      Top             =   900
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
      Left            =   8880
      Top             =   900
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
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   9120
      TabIndex        =   12
      Top             =   6540
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
      MICON           =   "frm_rpt_pajak.frx":24F0
      PICN            =   "frm_rpt_pajak.frx":250C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4965
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8758
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SSP"
      TabPicture(0)   =   "frm_rpt_pajak.frx":359E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "SPT MASA"
      TabPicture(1)   =   "frm_rpt_pajak.frx":35BA
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "Report Control Button"
         Height          =   1215
         Left            =   840
         TabIndex        =   15
         Top             =   2970
         Width           =   8415
         Begin prj_tpc.vbButton cmdPrintMonth 
            Height          =   705
            Left            =   6420
            TabIndex        =   16
            Top             =   300
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Print 1721"
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
            MICON           =   "frm_rpt_pajak.frx":35D6
            PICN            =   "frm_rpt_pajak.frx":35F2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdPrintDes 
            Height          =   705
            Left            =   6420
            TabIndex        =   17
            Top             =   300
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Print 1721"
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
            MICON           =   "frm_rpt_pajak.frx":4684
            PICN            =   "frm_rpt_pajak.frx":46A0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdPrint1721T 
            Height          =   705
            Left            =   5130
            TabIndex        =   18
            Top             =   300
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Print 1721-T"
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
            MICON           =   "frm_rpt_pajak.frx":5732
            PICN            =   "frm_rpt_pajak.frx":574E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdPrint1721II 
            Height          =   705
            Left            =   4110
            TabIndex        =   19
            Top             =   300
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Print 1721-II"
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
            MICON           =   "frm_rpt_pajak.frx":67E0
            PICN            =   "frm_rpt_pajak.frx":67FC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdPrint1721I 
            Height          =   705
            Left            =   3090
            TabIndex        =   20
            Top             =   300
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Print 1721-I"
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
            MICON           =   "frm_rpt_pajak.frx":788E
            PICN            =   "frm_rpt_pajak.frx":78AA
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
      Begin VB.Frame Frame3 
         Caption         =   "Report Control Button"
         Height          =   1215
         Left            =   -74160
         TabIndex        =   9
         Top             =   2970
         Width           =   8415
         Begin prj_tpc.vbButton cmdPrint 
            Height          =   705
            Left            =   7020
            TabIndex        =   10
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
            MICON           =   "frm_rpt_pajak.frx":893C
            PICN            =   "frm_rpt_pajak.frx":8958
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdPrintPreview 
            Height          =   705
            Left            =   5970
            TabIndex        =   11
            Top             =   300
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Preview"
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
            MICON           =   "frm_rpt_pajak.frx":99EA
            PICN            =   "frm_rpt_pajak.frx":9A06
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
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   3
         Top             =   270
         Width           =   8415
         Begin MSComCtl2.DTPicker DTPicker_ssp 
            Height          =   300
            Left            =   3510
            TabIndex        =   4
            Top             =   1140
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   83623939
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "MONTH"
            Height          =   195
            Left            =   2820
            TabIndex        =   5
            Top             =   1170
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   840
         TabIndex        =   2
         Top             =   270
         Width           =   8415
         Begin MSComCtl2.DTPicker DTPicker_masa 
            Height          =   300
            Left            =   3510
            TabIndex        =   13
            Top             =   1140
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   83623939
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "MONTH"
            Height          =   195
            Left            =   2820
            TabIndex        =   14
            Top             =   1170
            Width           =   600
         End
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT SSP"
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
      Left            =   330
      TabIndex        =   8
      Top             =   180
      Width           =   3285
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   930
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   -120
      Picture         =   "frm_rpt_pajak.frx":AA98
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12690
   End
End
Attribute VB_Name = "frm_rpt_pajak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim int_libur() As Integer

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrintPreview_Click()
Dim str_sql, str_param_periode, str_file As String
Dim a As New frm_rpt
    
        
    str_file = "\report\rpt_ssp.rpt"
    
    str_sql = "SELECT MONTH(CONCAT(a.month,'-01')) MONTH," & _
            "YEAR(CONCAT(a.month,'-01')) YEAR," & _
            "a.company_code,b.company_name,b.address," & _
            "b.npwp,b.city_name," & _
            "SUM(CASE WHEN salary_code='SU-36' THEN salary_value ELSE 0 END) AS pph21 " & _
          "FROM h_salary a JOIN m_company b ON a.company_code = b.company_code " & _
          "WHERE a.company_code = '" & TDBCombo_company.Text & "' " & _
            "AND a.month = '" & Format(DTPicker_ssp.Value, "yyyy-MM") & "' " & _
          "GROUP BY a.company_code"
    str_param_periode = "MONTHLY : (" & Format(d2, "yyyy-MM") & ")"
    
    Call a.Show
    a.Caption = "REPORT SSP"
    Call a.viewReport(str_sql, str_file, str_param_periode)
End Sub

Private Sub rpt_yearly()
Dim str_sql, str_param_periode, str_file As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim d1, d2 As String

    str_file = "\report\rpt_spt_pph21.rpt"
    
    int_flag_company = 1
    str_company_code = TDBCombo_company.Columns("company_code").Value
    int_flag_employee = 0
    
    If cbo_ssp_employee.Text = "All" Then
        str_employee_code = ""
    Else
        str_employee_code = txt_ssp_employee_code.Text
    End If
    
    d1 = Format(DTPicker_ssp.Value, "yyyy-01-01")
    d2 = Format(DTPicker_ssp.Value, "yyyy-12-27")
    
    str_sql = "CALL spr_pph21('" & d1 & "','" & d2 & "'," _
            & int_flag_company & ",'" & str_company_code & "','" & TDBCombo_ssp_department.Text & "'," _
            & int_flag_employee & ",'" & str_employee_code & "'," _
            & cbo_ssp_department.ListIndex & ",'" & LOGIN_LEVEL & "','" & LOGIN_CODE & "');"
    str_param_periode = "YEARLY : (" & Format(DTPicker_ssp.Value, "yyyy") & ")"
    
    Call a.Show
    a.Caption = "SPT PPH PASAL 21 REPORT"
    Call a.rpt_view_spt_pph21(str_sql, str_file, str_param_periode)
End Sub

Private Sub cmdPrint_Click()
Dim str_sql, str_param_periode, str_file As String
Dim a As New frm_rpt
    
        
    str_file = "\report\rpt_ssp_blank.rpt"
    
    str_sql = "SELECT MONTH(CONCAT(a.month,'-01')) MONTH," & _
            "YEAR(CONCAT(a.month,'-01')) YEAR," & _
            "a.company_code,b.company_name,b.address," & _
            "b.npwp,b.city_name," & _
            "SUM(CASE WHEN salary_code='SU-36' THEN salary_value ELSE 0 END) AS pph21 " & _
          "FROM h_salary a JOIN m_company b ON a.company_code = b.company_code " & _
          "WHERE a.company_code = '" & TDBCombo_company.Text & "' " & _
            "AND a.month = '" & Format(DTPicker_ssp.Value, "yyyy-MM") & "' " & _
          "GROUP BY a.company_code"
    str_param_periode = "MONTHLY : (" & Format(d2, "yyyy-MM") & ")"
    
    Call a.Show
    a.Caption = "REPORT SSP"
    Call a.rpt_print(str_sql, str_file, str_param_periode)
End Sub

Private Sub cmdPrintMonth_Click()
Dim str_sql, str_param_periode, str_file As String
Dim a As New frm_rpt
        
    str_file = "\report\rpt_spt_masa_pajak.rpt"
    
    str_sql = "CALL spr_masa_pajak('" & Format(DTPicker_masa.Value, "yyyy-MM") & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_LEVEL & "','" & LOGIN_CODE & "')"
    str_param_periode = "MONTHLY : (" & Format(d2, "yyyy-MM") & ")"
    
    Call a.Show
    a.Caption = "REPORT SPT MASA"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub cmdPrintDes_Click()
Dim str_sql, str_param_periode, str_file As String
Dim a As New frm_rpt
    
    str_file = "\report\rpt_spt_masa_pajak.rpt"
    
    str_sql = "CALL spr_masa_pajak_des('" & Format(DTPicker_masa.Value, "yyyy-MM") & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_LEVEL & "','" & LOGIN_CODE & "')"
    str_param_periode = "MONTHLY : (" & Format(d2, "yyyy-MM") & ")"
    
    Call a.Show
    a.Caption = "REPORT SPT MASA"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub cmdPrint1721T_Click()
Dim str_sql, str_param_periode, str_file As String
Dim a As New frm_rpt
        
    str_file = "\report\rpt_spt_masa_1721t.rpt"
    
    str_sql = "CALL spr_masa_1721t('" & Format(DTPicker_masa.Value, "yyyy-MM") & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_LEVEL & "','" & LOGIN_CODE & "')"
    str_param_periode = "MONTHLY : (" & Format(d2, "yyyy-MM") & ")"
    
    Call a.Show
    a.Caption = "REPORT SPT MASA 1721-T"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub cmdPrint1721II_Click()
Dim str_sql, str_param_periode, str_file As String
Dim a As New frm_rpt
        
    str_file = "\report\rpt_spt_masa_1721II.rpt"
    
    str_sql = "CALL spr_masa_1721t('" & Format(DTPicker_masa.Value, "yyyy-MM") & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_LEVEL & "','" & LOGIN_CODE & "')"
    str_param_periode = "MONTHLY : (" & Format(d2, "yyyy-MM") & ")"
    
    Call a.Show
    a.Caption = "REPORT SPT MASA 1721-II"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub cmdPrint1721I_Click()
Dim str_sql, str_param_periode, str_file As String
Dim a As New frm_rpt
        
    str_file = "\report\rpt_spt_masa_1721I.rpt"
    
    str_sql = "CALL spr_masa_1721I('" & Format(DTPicker_masa.Value, "yyyy-MM") & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_LEVEL & "','" & LOGIN_CODE & "')"
    str_param_periode = "MONTHLY : (" & Format(d2, "yyyy-MM") & ")"
    
    Call a.Show
    a.Caption = "REPORT SPT MASA 1721-I"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
End Sub

Private Sub DTPicker_masa_Change()
    If month(DTPicker_masa.Value) = 12 Then
        cmdPrintMonth.Visible = False
        cmdPrintDes.Visible = True
    Else
        cmdPrintMonth.Visible = True
        cmdPrintDes.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Adodc_company.ConnectionString = strConn
    
    Call load_data_company
    
    DTPicker_ssp.Value = Now
    
    Timer1.Enabled = True
    SSTab1.Tab = 0
End Sub

Private Sub load_data_company()
    Adodc_company.RecordSource = "select * from m_company order by company_code"
    Adodc_company.Refresh
    
    TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        DTPicker_ssp.Value = Now
    ElseIf SSTab1.Tab = 1 Then
        DTPicker_masa.Value = Now
        DTPicker_masa_Change
    End If
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call set_company_mode_adodc(Adodc_company, TDBCombo_company, txt_company_name)
End Sub



