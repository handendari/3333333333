VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_rpt_detail_attendance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LAPORAN - KEHADIRAN"
   ClientHeight    =   7350
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   10530
   Icon            =   "frm_rpt_att_detail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   Begin prj_panji.LynxGrid LynxGrid2 
      Height          =   2925
      Left            =   4680
      TabIndex        =   51
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
      Left            =   3300
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   40
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3300
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   840
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4155
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7329
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "DETAIL"
      TabPicture(0)   =   "frm_rpt_att_detail.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SUMMARY"
      TabPicture(1)   =   "frm_rpt_att_detail.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2865
         Left            =   -74160
         TabIndex        =   21
         Top             =   690
         Width           =   8415
         Begin VB.ComboBox cbo_sum_employee 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":05C2
            Left            =   1800
            List            =   "frm_rpt_att_detail.frx":05CC
            TabIndex        =   29
            Text            =   "..."
            Top             =   780
            Width           =   1695
         End
         Begin VB.Frame Frame4 
            Caption         =   "TYPE"
            Height          =   525
            Left            =   1800
            TabIndex        =   22
            Top             =   1140
            Width           =   4125
            Begin VB.OptionButton opt_sum_daily 
               Caption         =   "HARIAN"
               Height          =   255
               Left            =   180
               TabIndex        =   25
               Top             =   210
               Width           =   1005
            End
            Begin VB.OptionButton opt_sum_monthly 
               Caption         =   "BULANAN"
               Height          =   255
               Left            =   1200
               TabIndex        =   24
               Top             =   210
               Width           =   1245
            End
            Begin VB.OptionButton opt_sum_periode 
               Caption         =   "PERIODE"
               Height          =   255
               Left            =   2580
               TabIndex        =   23
               Top             =   210
               Width           =   1185
            End
         End
         Begin VB.Frame fra_sum_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   28
            Top             =   540
            Width           =   4575
            Begin VB.TextBox txt_sum_employee_code 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4170
               TabIndex        =   49
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.TextBox txt_sum_nik 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   0
               TabIndex        =   48
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox txt_sum_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               DragMode        =   1  'Automatic
               Height          =   285
               Left            =   1920
               TabIndex        =   47
               Top             =   240
               Width           =   2385
            End
            Begin prj_panji.vbButton cmd_sum_browse_employee 
               Height          =   285
               Left            =   1560
               TabIndex        =   50
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
               MICON           =   "frm_rpt_att_detail.frx":05DD
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
         Begin VB.Frame fra_sum_monthly 
            Height          =   585
            Left            =   1800
            TabIndex        =   30
            Top             =   1620
            Width           =   1875
            Begin MSComCtl2.DTPicker DTPicker_sum_monthly 
               Height          =   300
               Left            =   90
               TabIndex        =   31
               Top             =   180
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM"
               Format          =   86835203
               UpDown          =   -1  'True
               CurrentDate     =   39278
            End
         End
         Begin VB.Frame fra_sum_daily 
            Height          =   585
            Left            =   1800
            TabIndex        =   26
            Top             =   1620
            Width           =   1875
            Begin MSComCtl2.DTPicker DTPicker_sum_daily 
               Height          =   300
               Left            =   90
               TabIndex        =   27
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   86835203
               CurrentDate     =   39278
            End
         End
         Begin VB.Frame fra_sum_periode 
            Height          =   585
            Left            =   1800
            TabIndex        =   32
            Top             =   1620
            Width           =   4125
            Begin MSComCtl2.DTPicker DTPicker_sum_periode_from 
               Height          =   300
               Left            =   150
               TabIndex        =   33
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   86835203
               CurrentDate     =   39278
            End
            Begin MSComCtl2.DTPicker DTPicker_sum_periode_to 
               Height          =   300
               Left            =   2280
               TabIndex        =   34
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   86835203
               CurrentDate     =   39278
            End
            Begin VB.Label Label1 
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
               Height          =   255
               Left            =   1950
               TabIndex        =   35
               Top             =   210
               Width           =   285
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "KARYAWAN"
            Height          =   195
            Left            =   720
            TabIndex        =   36
            Top             =   840
            Width           =   930
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2835
         Left            =   840
         TabIndex        =   6
         Top             =   660
         Width           =   8415
         Begin VB.Frame Frame6 
            Caption         =   "TYPE"
            Height          =   525
            Left            =   1800
            TabIndex        =   17
            Top             =   1170
            Width           =   4125
            Begin VB.OptionButton opt_dtl_periode 
               Caption         =   "PERIODE"
               Height          =   255
               Left            =   2580
               TabIndex        =   20
               Top             =   210
               Width           =   1185
            End
            Begin VB.OptionButton opt_dtl_monthly 
               Caption         =   "BULANAN"
               Height          =   255
               Left            =   1200
               TabIndex        =   19
               Top             =   210
               Width           =   1245
            End
            Begin VB.OptionButton opt_dtl_daily 
               Caption         =   "HARIAN"
               Height          =   255
               Left            =   180
               TabIndex        =   18
               Top             =   210
               Width           =   1005
            End
         End
         Begin VB.Frame fra_dtl_employee 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   615
            Left            =   3600
            TabIndex        =   7
            Top             =   570
            Visible         =   0   'False
            Width           =   4575
            Begin VB.TextBox txt_dtl_employee_code 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4170
               TabIndex        =   45
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.TextBox txt_dtl_nik 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   0
               TabIndex        =   44
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox txt_dtl_employee_name 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               DragMode        =   1  'Automatic
               Height          =   285
               Left            =   1920
               TabIndex        =   43
               Top             =   240
               Width           =   2385
            End
            Begin prj_panji.vbButton cmd_dtl_browse_employee 
               Height          =   285
               Left            =   1560
               TabIndex        =   46
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
         Begin VB.ComboBox cbo_dtl_employee 
            Height          =   315
            ItemData        =   "frm_rpt_att_detail.frx":0615
            Left            =   1800
            List            =   "frm_rpt_att_detail.frx":061F
            TabIndex        =   2
            Text            =   "..."
            Top             =   810
            Width           =   1695
         End
         Begin VB.Frame fra_dtl_monthly 
            Height          =   585
            Left            =   1800
            TabIndex        =   11
            Top             =   1650
            Width           =   1875
            Begin MSComCtl2.DTPicker DTPicker_dtl_monthly 
               Height          =   300
               Left            =   90
               TabIndex        =   12
               Top             =   180
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM"
               Format          =   86835203
               UpDown          =   -1  'True
               CurrentDate     =   39278
            End
         End
         Begin VB.Frame fra_dtl_daily 
            Height          =   585
            Left            =   1800
            TabIndex        =   9
            Top             =   1650
            Width           =   1875
            Begin MSComCtl2.DTPicker DTPicker_dtl_daily 
               Height          =   300
               Left            =   90
               TabIndex        =   10
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   86835203
               CurrentDate     =   39278
            End
         End
         Begin VB.Frame fra_dtl_periode 
            Height          =   585
            Left            =   1800
            TabIndex        =   13
            Top             =   1650
            Width           =   4125
            Begin MSComCtl2.DTPicker DTPicker_dtl_periode_from 
               Height          =   300
               Left            =   150
               TabIndex        =   14
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   86835203
               CurrentDate     =   39278
            End
            Begin MSComCtl2.DTPicker DTPicker_dtl_periode_to 
               Height          =   300
               Left            =   2280
               TabIndex        =   15
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   86835203
               CurrentDate     =   39278
            End
            Begin VB.Label Label6 
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
               Height          =   255
               Left            =   1950
               TabIndex        =   16
               Top             =   210
               Width           =   285
            End
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "KARYAWAN"
            Height          =   195
            Left            =   720
            TabIndex        =   8
            Top             =   870
            Width           =   930
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Control Button"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   5940
      Width           =   10095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   180
         Top             =   360
      End
      Begin prj_panji.vbButton cmdExit 
         Height          =   705
         Left            =   8190
         TabIndex        =   37
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
         MICON           =   "frm_rpt_att_detail.frx":0630
         PICN            =   "frm_rpt_att_detail.frx":064C
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
         Left            =   7170
         TabIndex        =   38
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
         MICON           =   "frm_rpt_att_detail.frx":16DE
         PICN            =   "frm_rpt_att_detail.frx":16FA
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
      Left            =   1500
      OleObjectBlob   =   "frm_rpt_att_detail.frx":278C
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_division 
      Height          =   375
      Left            =   1500
      OleObjectBlob   =   "frm_rpt_att_detail.frx":46F2
      TabIndex        =   41
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DIVISI"
      Height          =   195
      Left            =   900
      TabIndex        =   42
      Top             =   1260
      Width           =   465
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "LAPORAN KEHADIRAN"
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
      TabIndex        =   39
      Top             =   150
      Width           =   4365
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   255
      TabIndex        =   5
      Top             =   870
      Width           =   1110
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_rpt_att_detail.frx":6659
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14850
   End
End
Attribute VB_Name = "frm_rpt_detail_attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCompany As New ADODB.Recordset
Dim rsDiv As New ADODB.Recordset

Dim int_libur() As Integer
Dim v_access As String
Dim v_dept As String

Private Sub rpt_detail_att()
Dim tgl1 As String, tgl2 As String, SQL As String
Dim str_param_periode As String
Dim a As New frm_rpt

    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
'    v_access = IIf(LOGIN_LEVEL = 100, "''", "(level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND flag_active = 0 order by employee_name")
'    v_dept = IIf(tdbcombo_division.Text = "", "company_code = '" & TDBCombo_company.Text & "'", "company_code = '" & TDBCombo_company.Text & "' AND department_code = '" & tdbcombo_division.Text & "'")
    
    If check_valid_periode Then
        If opt_dtl_daily.Value = True Then
            tgl1 = Format(DTPicker_dtl_daily.Value, "yyyy-MM-dd")
            tgl2 = Format(DTPicker_dtl_daily.Value, "yyyy-MM-dd")
        ElseIf opt_dtl_monthly.Value = True Then
            tgl1 = Format(DTPicker_dtl_monthly.Value, "yyyy-MM") & "-01"
            tgl2 = Format(DTPicker_dtl_monthly.Value, "yyyy-MM") & "-" & getEndDay(DTPicker_dtl_monthly.month, DTPicker_dtl_monthly.year)
        Else
            tgl1 = Format(DTPicker_dtl_periode_from.Value, "yyyy-MM-dd")
            tgl2 = Format(DTPicker_dtl_periode_to.Value, "yyyy-MM-dd")
        End If
            
        Call days_func(tgl1, tgl2)
            
        str_param_periode = Format(tgl1, "dd-MMM-yyyy") & " s/d " & Format(tgl2, "dd-MMM-yyyy")
        
        SQL = "call sp_detail_attendance('" & txt_dtl_employee_code.Text & "','" & tgl1 & "','" & tgl2 & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_CODE & "','" & TDBCombo_division.Text & "','" & IIf(LOGIN_LEVEL = 100, LOGIN_FULLNAME, EMPLOYEE_NAME) & "','" & LOGIN_LEVEL & "')"
        str_file = "\report\rpt_detail_attendance.rpt"
        
        Call a.Show
        a.Caption = "DETAIL ATTENDANCE"
         
        Call a.rpt_view(SQL, str_file, str_param_periode)
    End If

End Sub

Private Sub rpt_detail_sum()
Dim tgl1 As String, tgl2 As String, SQL As String
Dim str_param_periode As String
Dim a As New frm_rpt

    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
'    v_access = IIf(LOGIN_LEVEL = 100, "''", "(level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND flag_active = 0 order by employee_name")
'    v_dept = IIf(tdbcombo_division.Text = "", "company_code = '" & TDBCombo_company.Text & "'", "company_code = '" & TDBCombo_company.Text & "' AND department_code = '" & tdbcombo_division.Text & "'")
    
    If check_valid_periode Then
        If opt_sum_daily.Value = True Then
            tgl1 = Format(DTPicker_sum_daily.Value, "yyyy-MM-dd")
            tgl2 = Format(DTPicker_sum_daily.Value, "yyyy-MM-dd")
        ElseIf opt_sum_monthly.Value = True Then
            tgl1 = Format(DTPicker_sum_monthly.Value, "yyyy-MM") & "-01"
            tgl2 = Format(DTPicker_sum_monthly.Value, "yyyy-MM") & "-" & getEndDay(DTPicker_dtl_monthly.month, DTPicker_dtl_monthly.year)
        Else
            tgl1 = Format(DTPicker_sum_periode_from.Value, "yyyy-MM-dd")
            tgl2 = Format(DTPicker_sum_periode_to.Value, "yyyy-MM-dd")
        End If
        
        str_param_periode = Format(tgl1, "dd-MMM-yyyy") & " s/d " & Format(tgl2, "dd-MMM-yyyy")
        
        SQL = "call spr_summary_attendance('" & txt_sum_employee_code.Text & "','" & tgl1 & "','" & tgl2 & "'," & _
                "'" & TDBCombo_company.Text & "','" & LOGIN_LEVEL & "','" & TDBCombo_division.Text & "','" & IIf(LOGIN_LEVEL = 100, LOGIN_FULLNAME, EMPLOYEE_NAME) & "')"
        str_file = "\report\rpt_summary_attendance.rpt"
        
        Call a.Show
        a.Caption = "SUMMARY ATTENDANCE"
         
        Call a.rpt_view(SQL, str_file, str_param_periode)
    End If

End Sub

Private Sub cbo_dtl_employee_Click()
    If cbo_dtl_employee.ListIndex = 0 Then
        fra_dtl_employee.Visible = False
    Else
        fra_dtl_employee.Visible = True
    End If
    
    txt_dtl_employee_code = "": txt_dtl_nik = "": txt_dtl_employee_name = ""
End Sub

Private Sub cbo_sum_employee_Click()
    If cbo_sum_employee.ListIndex = 0 Then
        fra_sum_employee.Visible = False
    Else
        fra_sum_employee.Visible = True
    End If
    
    txt_sum_employee_code = "": txt_sum_nik = "": txt_sum_employee_name = ""
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Function check_valid_periode() As Boolean
check_valid_periode = True

    If SSTab1.Tab = 0 Then
        If cbo_dtl_employee.ListIndex = 1 And Trim(txt_dtl_nik.Text) = "" Then
            MsgBox "Karyawan Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            cmd_dtl_browse_employee.SetFocus
            check_valid_periode = False
            Exit Function
        End If
    Else
        If cbo_sum_employee.ListIndex = 1 And Trim(txt_sum_nik.Text) = "" Then
            MsgBox "Karyawan Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            cmd_sum_browse_employee.SetFocus
            check_valid_periode = False
            Exit Function
        End If
    End If
End Function

Private Sub cmdPrint_Click()
    If check_validate_tdbcombo(TDBCombo_company) = False Then
        MsgBox "Perusahaan Belum Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
    If SSTab1.Tab = 0 Then
        Call rpt_detail_att
    Else
        Call rpt_detail_sum
    End If
    
End Sub

Private Sub Form_Load()
    Call load_data_company
    Call load_data_user_access(Me)
    
    SSTab1.Tab = 0
    opt_dtl_daily.Value = True
    
    cbo_dtl_employee.ListIndex = 0
    cbo_sum_employee.ListIndex = 0
    
    Call createGridKar
    
    timer1.Enabled = True
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub opt_dtl_daily_Click()
    fra_dtl_daily.Visible = True
    fra_dtl_monthly.Visible = False
    fra_dtl_periode.Visible = False
    
    DTPicker_dtl_daily.Value = Now
End Sub

Private Sub opt_dtl_monthly_Click()
    fra_dtl_daily.Visible = False
    fra_dtl_monthly.Visible = True
    fra_dtl_periode.Visible = False
    
    DTPicker_dtl_monthly.Value = Now
End Sub

Private Sub opt_dtl_periode_Click()
    fra_dtl_daily.Visible = False
    fra_dtl_monthly.Visible = False
    fra_dtl_periode.Visible = True
    
    DTPicker_dtl_periode_from.Value = Now
    DTPicker_dtl_periode_to.Value = Now
End Sub

Private Sub opt_sum_daily_Click()
    fra_sum_daily.Visible = True
    fra_sum_monthly.Visible = False
    fra_sum_periode.Visible = False
    
    DTPicker_sum_daily.Value = Now
End Sub

Private Sub opt_sum_monthly_Click()
    fra_sum_daily.Visible = False
    fra_sum_monthly.Visible = True
    fra_sum_periode.Visible = False
    
    DTPicker_sum_monthly.Value = Now
End Sub

Private Sub opt_sum_periode_Click()
    fra_sum_daily.Visible = False
    fra_sum_monthly.Visible = False
    fra_sum_periode.Visible = True
    
    DTPicker_sum_periode_from.Value = Now
    DTPicker_sum_periode_to.Value = Now
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        opt_dtl_daily.Value = True
    Else
        opt_sum_daily.Value = True
    End If
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If
    
    Call load_data_division
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
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

Private Sub days_func(start_time As String, end_time As String)
Dim v_tgl_awal, v_tgl_akhir As Date

    v_tgl_awal = Format(start_time, "yyyy-MM-dd")
    v_tgl_akhir = Format(end_time, "yyyy-MM-dd")
    
    v_tgl_awal = DateValue(v_tgl_awal)
    v_tgl_akhir = DateValue(v_tgl_akhir)
    
    SQL = "delete from m_days where dt between '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' and '" & Format(v_tgl_akhir, "yyyy-MM-dd") & "'"
    CnG.Execute SQL
            
        While v_tgl_awal <= v_tgl_akhir
            SQL = "SELECT holiday_date FROM t_holiday WHERE date(holiday_date) = '" & Format(v_tgl_awal, "yyyy-MM-dd") & "'"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                SQL = "INSERT INTO m_days (dt,status,description) " & _
                      "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','L','HOLIDAY')"
                CnG.Execute SQL
            Else
                If Format(v_tgl_awal, "dddd") = "Sunday" Then
                    SQL = "INSERT INTO m_days (dt,status,description) " & _
                          "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','M','SUNDAY')"
                    CnG.Execute SQL
                ElseIf Format(v_tgl_awal, "dddd") = "Saturday" Then
                    SQL = "INSERT INTO m_days (dt,status,description) " & _
                          "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','S','SATURDAY')"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_days (dt,status,description) " & _
                          "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','W','WORK DAY')"
                    CnG.Execute SQL
                End If
            End If
            rs.Close
            
            
            v_tgl_awal = v_tgl_awal + 1
        Wend
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
            If SSTab1.Tab = 0 Then
                SQL = "SELECT a.nik,a.employee_name," _
                            & "a.division_code,b.division_name," _
                            & "a.title_code,c.title_name,a.employee_code " _
                        & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                        & "JOIN m_title c ON a.title_code = c.title_code " _
                        & "JOIN m_company e ON a.company_code = e.company_code " _
                        & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                                "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                            & "AND (a.nik LIKE '%" & txt_dtl_nik.Text & "%' " _
                            & "OR a.employee_name LIKE '%" & txt_dtl_nik.Text & "%') " _
                            & "AND a.flag_active <> 0"
            Else
                SQL = "SELECT a.nik,a.employee_name," _
                            & "a.division_code,b.division_name," _
                            & "a.title_code,c.title_name,a.employee_code " _
                        & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                        & "JOIN m_title c ON a.title_code = c.title_code " _
                        & "JOIN m_company e ON a.company_code = e.company_code " _
                        & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                                "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                            & "AND (a.nik LIKE '%" & txt_sum_nik.Text & "%' " _
                            & "OR a.employee_name LIKE '%" & txt_sum_nik.Text & "%') " _
                            & "AND a.flag_active <> 0"
            End If
        Else
            If SSTab1.Tab = 0 Then
                SQL = "SELECT a.nik,a.employee_name," _
                            & "a.division_code,b.division_name," _
                            & "a.title_code,c.title_name,a.employee_code " _
                        & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                        & "JOIN m_title c ON a.title_code = c.title_code " _
                        & "JOIN m_company e ON a.company_code = e.company_code " _
                        & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                                "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                            & "AND (a.nik LIKE '%" & txt_dtl_nik.Text & "%' " _
                            & "OR a.employee_name LIKE '%" & txt_dtl_nik.Text & "%') " _
                            & "AND a.flag_active <> 0 AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                            & "ORDER BY a.employee_name ASC"
            Else
                SQL = "SELECT a.nik,a.employee_name," _
                            & "a.division_code,b.division_name," _
                            & "a.title_code,c.title_name,a.employee_code " _
                        & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                        & "JOIN m_title c ON a.title_code = c.title_code " _
                        & "JOIN m_company e ON a.company_code = e.company_code " _
                        & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                                "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                            & "AND (a.nik LIKE '%" & txt_sum_nik.Text & "%' " _
                            & "OR a.employee_name LIKE '%" & txt_sum_nik.Text & "%') " _
                            & "AND a.flag_active <> 0 AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                            & "ORDER BY a.employee_name ASC"

            End If
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
                
                If SSTab1.Tab = 0 Then
                    txt_dtl_employee_code.Text = rs!employee_code
                    txt_dtl_employee_name.Text = rs!EMPLOYEE_NAME
                    txt_dtl_nik.Text = rs!nik
                Else
                    txt_sum_employee_code.Text = rs!employee_code
                    txt_sum_employee_name.Text = rs!EMPLOYEE_NAME
                    txt_sum_nik.Text = rs!nik
                End If
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
            If SSTab1.Tab = 0 Then
                txt_dtl_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
                txt_dtl_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
                txt_dtl_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
            Else
                txt_sum_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
                txt_sum_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
                txt_sum_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
            End If
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

Private Sub txt_dtl_nik_Change()
    If txt_dtl_nik.Text = "" Then
        txt_dtl_employee_code.Text = ""
        txt_dtl_employee_name.Text = ""
    End If
End Sub

Private Sub txt_sum_nik_Change()
    If txt_sum_nik.Text = "" Then
        txt_sum_employee_code.Text = ""
        txt_sum_employee_name.Text = ""
    End If
End Sub

Private Sub txt_dtl_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub txt_sum_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmd_dtl_browse_employee_Click()
    isiGridKar (1)
End Sub

Private Sub cmd_sum_browse_employee_Click()
    isiGridKar (1)
End Sub

