VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdi_absensi 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000A&
   Caption         =   "e-FREIGHT ATTENDANCE RELEASE 1.5 VERSION"
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11295
   Icon            =   "mdi_absensi.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer timer_get_log_data 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   240
      Top             =   1680
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_absensi.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_absensi.frx":09DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_absensi.frx":0F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_absensi.frx":1510
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_absensi.frx":182A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      ToolTips        =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "key_setting"
            Object.ToolTipText     =   "Setting ..."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "key_employee"
            Object.ToolTipText     =   "Employee ..."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "key_detail_report"
            Object.ToolTipText     =   "Detail Report..."
            Object.Tag             =   "Detail Report..."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "key_about"
            Object.ToolTipText     =   "About ..."
            ImageIndex      =   5
         EndProperty
      EndProperty
      MouseIcon       =   "mdi_absensi.frx":1F7C
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   6690
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2647
            MinWidth        =   2647
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6704
            MinWidth        =   6704
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "mdi_absensi.frx":2516
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnu_setting 
         Caption         =   "Setting"
      End
      Begin VB.Menu mnu_change_login 
         Caption         =   "Change Login"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&Master"
      Begin VB.Menu mnu_mst_device 
         Caption         =   "Device"
      End
      Begin VB.Menu mnu_mst_enroll 
         Caption         =   "&Enroll"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu_mst_company 
         Caption         =   "Company"
      End
      Begin VB.Menu mnu_mst_department 
         Caption         =   "Department"
      End
      Begin VB.Menu mnu_mst_division 
         Caption         =   "Division"
      End
      Begin VB.Menu mnu_mst_title 
         Caption         =   "Title"
      End
      Begin VB.Menu mnu_mst_employee 
         Caption         =   "&Employee"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnu_mst_working_time 
         Caption         =   "Working Time"
         Begin VB.Menu mnu_mst_general 
            Caption         =   "General"
         End
         Begin VB.Menu mnu_mst_shift 
            Caption         =   "Shift"
         End
      End
      Begin VB.Menu mnu_mst_day 
         Caption         =   "Working Day"
         Begin VB.Menu mnu_mst_wd_general 
            Caption         =   "General"
         End
         Begin VB.Menu mnu_mst_wd_shift 
            Caption         =   "Shift"
         End
      End
      Begin VB.Menu mnu_mst_user 
         Caption         =   "User"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "&Utility"
      Begin VB.Menu mnu_trans_working_time 
         Caption         =   "Employees Working Time"
      End
      Begin VB.Menu mnu_trans_log_att 
         Caption         =   "Log Attendance"
      End
      Begin VB.Menu mnu_trans_manual_check 
         Caption         =   "Manual Check"
      End
      Begin VB.Menu mnu_trans_absent 
         Caption         =   "Absent"
      End
      Begin VB.Menu mnu_trans_leave_sub 
         Caption         =   "Leave"
         Begin VB.Menu mnu_trans_leave 
            Caption         =   "Employee Leave"
         End
         Begin VB.Menu mnu_trans_general_leave 
            Caption         =   "General Leave"
         End
         Begin VB.Menu mnu_trans_summary_leave 
            Caption         =   "Summary Leave"
         End
      End
      Begin VB.Menu mnu_trans_duty 
         Caption         =   "Duty"
      End
      Begin VB.Menu mnu_trans_holiday 
         Caption         =   "Holiday"
      End
      Begin VB.Menu mnu_trans_employee_performance 
         Caption         =   "Employee Performance"
      End
   End
   Begin VB.Menu mnu_report 
      Caption         =   "&Report"
      Begin VB.Menu mnu_rpt_master 
         Caption         =   "Master Data"
      End
      Begin VB.Menu mnu_rpt_detail_attendance 
         Caption         =   "Detail Attendance"
      End
      Begin VB.Menu mnu_rpt_summary_attendance 
         Caption         =   "Summary Attendance"
      End
      Begin VB.Menu mnu_rpt_summary_leave 
         Caption         =   "Summary Leave"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "mdi_absensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str_form_title(30) As String


Private Sub mnu_RptPembelian_Click()
If frmRptPembelian.WindowState = 1 Then frmRptPembelian.WindowState = 0
frmRptPembelian.Top = 0
frmRptPembelian.Left = 0
frmRptPembelian.Width = 11880
frmRptPembelian.Height = 6240
frmRptPembelian.Show
frmRptPembelian.SetFocus
End Sub

Private Sub mnu_mst_costcenter_Click()
If frm_mst_costcenter.WindowState = 1 Then frm_mst_costcenter.WindowState = 0
frm_mst_costcenter.Top = 0
frm_mst_costcenter.Left = 0
frm_mst_costcenter.Width = 11880
frm_mst_costcenter.Height = 4665
frm_mst_costcenter.Show
frm_mst_costcenter.SetFocus
End Sub

Private Sub mnu_mst_kategory_Click()
If frm_mst_kategory.WindowState = 1 Then frm_mst_kategory.WindowState = 0
frm_mst_kategory.Top = 0
frm_mst_kategory.Left = 0
frm_mst_kategory.Width = 9165
frm_mst_kategory.Height = 5085
frm_mst_kategory.Show
frm_mst_kategory.SetFocus
End Sub

Private Sub Command1_Click()
MsgBox check_auto_log
End Sub

Private Sub MDIForm_Load()
With StatusBar1
    .Panels(1).Text = " DATE : " & UCase(Format(Date, "dddd, dd MMMM yyyy"))
    .Panels(4).Text = "LOGIN NAME  : " & UCase(LOGIN_NAME)
End With

Call load_all_menu
Call disable_all_menu
Call load_user_menu
BLN_RUNNING = True
Call get_setting_device
Call get_setting_auto_log
timer_get_log_data.Enabled = BLN_AUTO_LOG

'Form1.Show 1
End Sub

Private Sub mnu_export_enroll_id_Click()
frm_trans_exp_enroll.Show
End Sub

Private Sub mnu_change_login_Click()
Call frm_etc_login.Form_Load
frm_etc_login.Show
End Sub

Private Sub mnu_mst_company_Click()
With frm_mst_company
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 14805
    .Height = 7200
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_department_Click()
With frm_mst_department
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 11880
    .Height = 7200
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_device_Click()
With frm_mst_device
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 7650
    .Height = 4545
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_division_Click()
With frm_mst_division
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 11880
    .Height = 7200
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_employee_Click()
With frm_mst_employee
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 14805
    .Height = 9600
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_enroll_Click()
With frm_mst_enroll
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 14805
    .Height = 9600
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_karyawan_Click()
With frm_mst_karyawan
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 14805
    .Height = 9495
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_rekening_Click()
If frm_mst_rekening.WindowState = 1 Then frm_mst_rekening.WindowState = 0
frm_mst_rekening.Top = 0
frm_mst_rekening.Left = 0
frm_mst_rekening.Width = 14670 '11880
frm_mst_rekening.Height = 9555 '7590
frm_mst_rekening.Show
frm_mst_rekening.WindowState = vbNormal
frm_mst_rekening.SetFocus
End Sub

Private Sub mnu_RptJurnal_Click()
If frmRptJurnal.WindowState = 1 Then frmRptJurnal.WindowState = 0
frmRptJurnal.Top = 0
frmRptJurnal.Left = 0
frmRptJurnal.Width = 11880
frmRptJurnal.Height = 6240
frmRptJurnal.Show
frmRptJurnal.SetFocus
End Sub

Private Sub mnu_stg_cost_center_Click()
If frm_stg_cost_center.WindowState = 1 Then frm_stg_cost_center.WindowState = 0
frm_stg_cost_center.Top = 0
frm_stg_cost_center.Left = 0
frm_stg_cost_center.Width = 7995
frm_stg_cost_center.Height = 4230
frm_stg_cost_center.Show
frm_stg_cost_center.SetFocus
End Sub

Private Sub mnu_stg_format_kode_rekening_Click()
If frm_stg_format_rekening.WindowState = 1 Then frm_stg_format_rekening.WindowState = 0
frm_stg_format_rekening.Top = 0
frm_stg_format_rekening.Left = 0
frm_stg_format_rekening.Width = 7995
frm_stg_format_rekening.Height = 4230
frm_stg_format_rekening.Show
frm_stg_format_rekening.SetFocus
End Sub

Private Sub mnu_stg_periode_aktif_Click()
If frm_stg_periode_aktif.WindowState = 1 Then frm_stg_periode_aktif.WindowState = 0
frm_stg_periode_aktif.Top = 0
frm_stg_periode_aktif.Left = 0
frm_stg_periode_aktif.Width = 7500
frm_stg_periode_aktif.Height = 4950
frm_stg_periode_aktif.Show
frm_stg_periode_aktif.SetFocus
End Sub

Private Sub mnu_stg_rekening_per_cost_center_Click()
If frm_stg_rek_cost_center.WindowState = 1 Then frm_stg_rek_cost_center.WindowState = 0
frm_stg_rek_cost_center.Top = 0
frm_stg_rek_cost_center.Left = 0
frm_stg_rek_cost_center.Width = 14580
frm_stg_rek_cost_center.Height = 9555
frm_stg_rek_cost_center.Show
frm_stg_rek_cost_center.SetFocus
End Sub

Private Sub mnu_stg_user_Click()
If frm_stg_user.WindowState = 1 Then frm_stg_user.WindowState = 0
frm_stg_user.Top = 0
frm_stg_user.Left = 0
frm_stg_user.Width = 13890
frm_stg_user.Height = 8940
frm_stg_user.Show
frm_stg_user.SetFocus
End Sub

Private Sub mnu_mst_general_Click()
With frm_mst_general
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 12735
    .Height = 6165
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_shift_Click()
With frm_mst_shift
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 12735
    .Height = 6165
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_title_Click()
With frm_mst_title
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 13890
    .Height = 9045
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_user_Click()
With frm_mst_user
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 13890
    .Height = 9045
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_wd_general_Click()
With frm_mst_working_day
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 12735
    .Height = 6165
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_mst_wd_shift_Click()
With frm_mst_working_day_shift
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 12735
    .Height = 6165
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_rpt_kehadiran_Click()
With frm_rpt_kehadiran
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 8970
    .Height = 5595
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_stg_device_Click()
frm_stg_all.Show
End Sub

Private Sub mnu_stg_shift_Click()
With frm_trans_shift
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 14805
    .Height = 9495
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_stg_timer_log_data_Click()
frm_stg_timer_log_data.Show
End Sub

Private Sub mnu_rpt_detail_attendance_Click()
With frm_rpt_att_detail
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 10680
    .Height = 7080
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_rpt_master_Click()
With frm_rpt_master
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 10680
    .Height = 7080
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_rpt_summary_attendance_Click()
With frm_rpt_att_summary
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 10680
    .Height = 7080
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_rpt_summary_leave_Click()
With frm_rpt_leave_summary
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 10680
    .Height = 7080
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_setting_Click()
With frm_stg_all
    If .WindowState = 1 Then .WindowState = 0
'    .Top = 0
'    .Left = 0
'    .Width = 14805
'    .Height = 9600
    .Show 1
'    .SetFocus
End With
End Sub

Private Sub mnu_trans_absent_Click()
With frm_trans_absent
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 14805
    .Height = 9600
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_trans_duty_Click()
With frm_trans_duty
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 14805
    .Height = 9600
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_trans_employee_performance_Click()
With frm_trans_performance
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 11880
    .Height = 7200
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_trans_general_leave_Click()
With frm_trans_general_leave
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 10215
    .Height = 9600
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_trans_holiday_Click()
With frm_trans_holiday
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 10215
    .Height = 9600
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_trans_leave_Click()
With frm_trans_leave
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 14805
    .Height = 9600
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_trans_log_att_Click()
With frm_trans_log_attendance
    If .WindowState = 1 Then .WindowState = 0
    If .Visible = False Then .Visible = True
    .Top = 0
    .Left = 0
    .Width = 14805
    .Height = 9495
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_trans_manual_check_Click()
With frm_trans_manual_check
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 14805
    .Height = 9600
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_trans_summary_leave_Click()
With frm_trans_leave_periode
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 13335
    .Height = 9600
    .Show
    .SetFocus
End With
End Sub

Private Sub mnu_trans_working_time_Click()
With frm_trans_employees_wt
    If .WindowState = 1 Then .WindowState = 0
    .Top = 0
    .Left = 0
    .Width = 14805
    .Height = 9600
    .Show
    .SetFocus
End With
End Sub

Private Sub mnuAbout_Click()
If frm_etc_about.WindowState = 1 Then frm_etc_about.WindowState = 0
frm_etc_about.Top = 0
frm_etc_about.Left = 0
frm_etc_about.Width = 7110
frm_etc_about.Height = 4290
frm_etc_about.Show
frm_etc_about.SetFocus
End Sub

Private Sub mnuBackupData_Click()
frmBackup.Show 1
End Sub

Private Sub mnuBukuBesar_Click()
If frmRptBukuBesar.WindowState = 1 Then frmRptBukuBesar.WindowState = 0
frmRptBukuBesar.Top = 0
frmRptBukuBesar.Left = 0
frmRptBukuBesar.Width = 11880
frmRptBukuBesar.Height = 6240
frmRptBukuBesar.Show
frmRptBukuBesar.SetFocus
End Sub

Private Sub mnuCostCenter_Click()

End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuKas_Click()
If frmJurnalKas.WindowState = 1 Then frmJurnalKas.WindowState = 0
frmJurnalKas.Top = 0
frmJurnalKas.Left = 0
frmJurnalKas.Width = 11880
frmJurnalKas.Height = 7590
frmJurnalKas.Show
frmJurnalKas.SetFocus
End Sub

Private Sub mnuJurnalMemorial_Click()
If frm_trans_memorial.WindowState = 1 Then frm_trans_memorial.WindowState = 0
frm_trans_memorial.Top = 0
frm_trans_memorial.Left = 0
frm_trans_memorial.Width = 14800
frm_trans_memorial.Height = 9500
frm_trans_memorial.Show
frm_trans_memorial.SetFocus
End Sub

Private Sub mnuMataUang_Click()
If frmMataUang.WindowState = 1 Then frmMataUang.WindowState = 0
frmMataUang.Top = 0
frmMataUang.Left = 0
frmMataUang.Width = 11880
frmMataUang.Show
frmMataUang.SetFocus
End Sub

Private Sub mnuRekening_B_Click()
If frmRekeningCostCenter.WindowState = 1 Then frmRekeningCostCenter.WindowState = 0
frmRekeningCostCenter.Top = 0
frmRekeningCostCenter.Left = 0
frmRekeningCostCenter.Width = 11880
frmRekeningCostCenter.Height = 7590
frmRekeningCostCenter.Show
frmRekeningCostCenter.SetFocus
End Sub

Private Sub mnuKB_Keluar_Click()
If frmJurnalKasBank_Keluar.WindowState = 1 Then _
        frmJurnalKasBank_Keluar.WindowState = 0
frmJurnalKasBank_Keluar.Top = 0
frmJurnalKasBank_Keluar.Left = 0
frmJurnalKasBank_Keluar.Width = 11880
frmJurnalKasBank_Keluar.Height = 7590
frmJurnalKasBank_Keluar.Show
frmJurnalKasBank_Keluar.SetFocus
End Sub

Private Sub mnuKB_Masuk_Click()
If frmJurnalKasBank_Masuk.WindowState = 1 Then _
        frmJurnalKasBank_Masuk.WindowState = 0
frmJurnalKasBank_Masuk.Top = 0
frmJurnalKasBank_Masuk.Left = 0
frmJurnalKasBank_Masuk.Width = 11880
frmJurnalKasBank_Masuk.Height = 7590
frmJurnalKasBank_Masuk.Show
frmJurnalKasBank_Masuk.SetFocus
End Sub

Private Sub mnuJurnalPembelian_Click()
If frmJurnalPB.WindowState = 1 Then frmJurnalPB.WindowState = 0
frmJurnalPB.Top = 0
frmJurnalPB.Left = 0
frmJurnalPB.Width = 11880
frmJurnalPB.Height = 7590
frmJurnalPB.Show
frmJurnalPB.SetFocus
End Sub

Private Sub mnuKBIn_Click()
If frm_trans_kas_bank_in.WindowState = 1 Then frm_trans_kas_bank_in.WindowState = 0
frm_trans_kas_bank_in.Top = 0
frm_trans_kas_bank_in.Left = 0
frm_trans_kas_bank_in.Width = 11880
frm_trans_kas_bank_in.Height = 7590
frm_trans_kas_bank_in.Show
frm_trans_kas_bank_in.SetFocus
End Sub

Private Sub mnuKBOut_Click()
If frmJurnalKasBank_Out.WindowState = 1 Then frmJurnalKasBank_Out.WindowState = 0
frmJurnalKasBank_Out.Top = 0
frmJurnalKasBank_Out.Left = 0
frmJurnalKasBank_Out.Width = 11880
frmJurnalKasBank_Out.Height = 7590
frmJurnalKasBank_Out.Show
frmJurnalKasBank_Out.SetFocus
End Sub

Private Sub mnuModal_Click()
If frmRptModal.WindowState = 1 Then frmRptModal.WindowState = 0
frmRptModal.Top = 0
frmRptModal.Left = 0
frmRptModal.Width = 11880
frmRptModal.Height = 6240
frmRptModal.Show
frmRptModal.SetFocus
End Sub

Private Sub mnuMasterSupplier_Click()
If frmSupplier.WindowState = 1 Then frmSupplier.WindowState = 0
frmSupplier.Top = 0
frmSupplier.Left = 0
frmSupplier.Width = 11880
frmSupplier.Show
frmSupplier.SetFocus
End Sub

Private Sub mnuNeraca_Click()
If frmRptNeraca.WindowState = 1 Then frmRptNeraca.WindowState = 0
frmRptNeraca.Top = 0
frmRptNeraca.Left = 0
frmRptNeraca.Width = 11880
frmRptNeraca.Height = 6240
frmRptNeraca.Show
frmRptNeraca.SetFocus
End Sub

Private Sub mnuRekening_Click()

End Sub

Private Sub mnuReportJurnal_Click()
If frmRptJurnal.WindowState = 1 Then frmRptJurnal.WindowState = 0
frmRptJurnal.Top = 0
frmRptJurnal.Left = 0
frmRptJurnal.Width = 11880
frmRptJurnal.Height = 6240
frmRptJurnal.Show
frmRptJurnal.SetFocus
End Sub

Private Sub mnuSaldoAwal_Click()
If frm_trans_sa_rek_cost_center.WindowState = 1 Then frm_trans_sa_rek_cost_center.WindowState = 0
frm_trans_sa_rek_cost_center.Top = 0
frm_trans_sa_rek_cost_center.Left = 0
frm_trans_sa_rek_cost_center.Width = 14580
frm_trans_sa_rek_cost_center.Height = 9555
frm_trans_sa_rek_cost_center.Show
frm_trans_sa_rek_cost_center.SetFocus
End Sub

Private Sub mnuSettingDepartment_Click()
If frmSettingRekening.WindowState = 1 Then frmSettingRekening.WindowState = 0
frmSettingRekening.Top = 0
frmSettingRekening.Left = 0
frmSettingRekening.Width = 11880
frmSettingRekening.Height = 6240
frmSettingRekening.Show
frmSettingRekening.SetFocus
End Sub

Private Sub mnuSettingDepartment_Lokal_Click()

End Sub

Private Sub mnuSettingGeneral_FR_Click()

End Sub

Private Sub mnuSettingKodeJurnal_Click()
If frmSetting_KodeJurnal.WindowState = 1 Then frmSetting_KodeJurnal.WindowState = 0
frmSetting_KodeJurnal.Top = 0
frmSetting_KodeJurnal.Left = 0
frmSetting_KodeJurnal.Width = 11880
frmSetting_KodeJurnal.Height = 6240
frmSetting_KodeJurnal.Show
frmSetting_KodeJurnal.SetFocus
End Sub

Private Sub mnuSettingLain2_Click()

End Sub

Private Sub mnuSettingReport_Click()
If frmSettingReport.WindowState = 1 Then frmSettingReport.WindowState = 0
frmSettingReport.Top = 0
frmSettingReport.Left = 0
frmSettingReport.Width = 11880
frmSettingReport.Height = 6240
frmSettingReport.Show
frmSettingReport.SetFocus
End Sub

Private Sub mnuSettingUserAccount_Click()

End Sub

Private Sub mnuSubRekening_Click()
If frmSubRekening.WindowState = 1 Then frmSubRekening.WindowState = 0
frmSubRekening.Top = 0
frmSubRekening.Left = 0
frmSubRekening.Width = 11880
frmSubRekening.Height = 7590
frmSubRekening.Show
frmSubRekening.WindowState = vbNormal
frmSubRekening.SetFocus
End Sub

Private Sub mnuTutupBukuGeneral_Click()
If frmClosingGeneral.WindowState = 1 Then frmClosingGeneral.WindowState = 0
frmClosingGeneral.Top = 0
frmClosingGeneral.Left = 0
frmClosingGeneral.Width = 11880
frmClosingGeneral.Height = 6240
frmClosingGeneral.Show
frmClosingGeneral.SetFocus
End Sub

Private Sub timer_get_log_data_Timer()
int_timer_tick = int_timer_tick + 1

If check_auto_log Then
    Call mnu_trans_log_att_Click
    frm_trans_log_attendance.public_int_caller = 1
    Call frm_trans_log_attendance.cmd_download_Click
    Unload frm_trans_log_attendance
    int_timer_tick = 0
End If
End Sub

Private Function check_auto_log() As Boolean
Dim rs As New ADODB.Recordset
check_auto_log = False

rs.Open "select count(*) as rec_count from s_auto_log where left(cast(s_time as time),5)='" & Format(Now, "hh:mm") & "'", CnG, adOpenStatic, adLockReadOnly
check_auto_log = IIf(rs.Fields("rec_count").Value >= 1 And mnu_trans_log_att.Enabled = True, True, False)
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "KEY_SETTING"
        Call mnu_stg_device_Click
    Case "KEY_EMPLOYEE"
        Call mnu_mst_employee_Click
    Case "KEY_DETAIL_REPORT"
        Call mnu_rpt_detail_attendance_Click

    Case "KEY_ABOUT"
        Call mnuAbout_Click
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick _
(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case UCase(ButtonMenu.Key)
    Case "KEYKBIN"
        Call mnuKBIn_Click
    Case "KEYKBOUT"
        Call mnuKBOut_Click
    Case "KEYPEMBELIAN"
        Call mnuJurnalPembelian_Click
    Case "KEYMEMO"
        Call mnuJurnalMemorial_Click
    
    Case "KEYJURNALKAS/BANK"
        Call mnuReportJurnal_Click
    Case "KEYREPORTPEMBELIAN"
        Call mnu_RptPembelian_Click
    Case "KEYBUKUBESAR"
        Call mnuBukuBesar_Click
    Case "KEYNERACA"
        Call mnuNeraca_Click
    Case "KEYMODAL"
        Call mnuModal_Click
    
    Case "KEYSETTINGGENERAL"
        Call mnuSettingGeneral_FR_Click
    Case "KEYSETTINGLOCAL"
        Call mnuSettingDepartment_Lokal_Click
    Case "KEYSETTINGKODEJURNAL"
        Call mnuSettingKodeJurnal_Click
    Case "KEYSETTINGREKENING"
        Call mnuSettingDepartment_Click
    Case "KEYSETTINGREPORT"
        Call mnuSettingReport_Click
    Case "KEYLAIN2"
        Call mnuSettingLain2_Click
    Case "KEYSETTINGUSERACCOUNT"
        Call mnuSettingUserAccount_Click
End Select
End Sub

Private Sub load_user_menu()
Dim rs As New ADODB.Recordset
Dim str_sql As String

If LOGIN_LEVEL = 100 Then
    Call enable_all_menu
    Exit Sub
End If

str_sql = "select F.form_code, F.form_name, F.form_title, " _
    & "F.allow_read, F.allow_add, F.allow_edit, F.allow_delete, F.allow_posting " _
    & "from m_user U inner join t_user F on U.user_code = F.user_code " _
    & "inner join m_form M on F.form_code = M.form_code " _
    & "where U.user_code = " & LOGIN_ID & " and user_name = '" & LOGIN_NAME _
    & "' and user_pass = '" & LOGIN_PASS & "' and M.user_level <= " & LOGIN_LEVEL


rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        Call set_enable_menu(rs.Fields("form_title").Value, _
                rs.Fields("allow_read").Value)
        rs.MoveNext
    Wend
End If
End Sub

Private Sub load_all_menu()
str_form_title(1) = "MASTER COMPANY"
str_form_title(2) = "MASTER DEPARTMENT"
str_form_title(3) = "MASTER DIVISION"
str_form_title(4) = "MASTER TITLE"
str_form_title(5) = "MASTER EMPLOYEE"
str_form_title(6) = "MASTER WORKING TIME (GENERAL)"
str_form_title(7) = "MASTER WORKING TIME (SHIFT)"
str_form_title(8) = "MASTER WORKING DAY (GENERAL)"
str_form_title(9) = "MASTER WORKING DAY (SHIFT)"
str_form_title(10) = "MASTER USER"
str_form_title(11) = "EMPLOYEES WORKING TIME"
str_form_title(12) = "LOG ATTENDANCE"
str_form_title(13) = "MANUAL CHECK UTILITY"
str_form_title(14) = "ABSENT UTILITY"
str_form_title(15) = "LEAVE UTILITY"
str_form_title(16) = "DUTY UTILITY"
str_form_title(17) = "HOLIDAY UTILITY"
str_form_title(18) = "GENERAL LEAVE UTILITY"

str_form_title(19) = "DETAIL REPORT"
str_form_title(20) = "SETTING"
str_form_title(21) = "MASTER DEVICE"
str_form_title(22) = "MASTER ENROLL"
str_form_title(23) = "SUMMARY REPORT"
str_form_title(24) = "MASTER REPORT"

str_form_title(25) = "SUMMARY LEAVE"
str_form_title(26) = "SUMMARY LEAVE REPORT"
str_form_title(27) = "EMPLOYEE PERFORMANCE"
End Sub

Private Sub enable_all_menu()
Dim i As Integer
For i = 1 To 27
    Call set_enable_menu(str_form_title(i), True)
Next
End Sub

Private Sub disable_all_menu()
Dim i As Integer
For i = 1 To 25
    Call set_enable_menu(str_form_title(i), False)
Next
End Sub

Private Sub set_enable_menu(ByVal strCaption As String, ByVal blnEnable As Boolean)
Select Case UCase(strCaption)
    Case "MASTER COMPANY"
        mnu_mst_company.Enabled = blnEnable
    Case "MASTER DEPARTMENT"
        mnu_mst_department.Enabled = blnEnable
    Case "MASTER DIVISION"
        mnu_mst_division.Enabled = blnEnable
    Case "MASTER TITLE"
        mnu_mst_title.Enabled = blnEnable
    Case "MASTER EMPLOYEE"
        mnu_mst_employee.Enabled = blnEnable
        Toolbar1.Buttons(3).Enabled = blnEnable
    
    Case "MASTER WORKING TIME (GENERAL)"
        mnu_mst_general.Enabled = blnEnable
    Case "MASTER WORKING TIME (SHIFT)"
        mnu_mst_shift.Enabled = blnEnable
    Case "MASTER WORKING DAY (GENERAL)"
        mnu_mst_wd_general.Enabled = blnEnable
    Case "MASTER WORKING DAY (SHIFT)"
        mnu_mst_wd_shift.Enabled = blnEnable
    Case "MASTER USER"
        mnu_mst_user.Enabled = blnEnable
        
    Case "EMPLOYEES WORKING TIME"
        mnu_trans_working_time.Enabled = blnEnable
    Case "LOG ATTENDANCE"
        mnu_trans_log_att.Enabled = blnEnable
    Case "MANUAL CHECK UTILITY"
        mnu_trans_manual_check.Enabled = blnEnable
    Case "ABSENT UTILITY"
        mnu_trans_absent.Enabled = blnEnable
    Case "LEAVE UTILITY"
        mnu_trans_leave.Enabled = blnEnable
    Case "DUTY UTILITY"
        mnu_trans_duty.Enabled = blnEnable
    Case "HOLIDAY UTILITY"
        mnu_trans_holiday.Enabled = blnEnable
    Case "GENERAL LEAVE UTILITY"
        mnu_trans_general_leave.Enabled = blnEnable
        
    Case "DETAIL REPORT"
        mnu_rpt_detail_attendance.Enabled = blnEnable
        Toolbar1.Buttons(5).Enabled = blnEnable
        
    Case "SETTING"
        mnu_setting.Enabled = blnEnable
        Toolbar1.Buttons(1).Enabled = blnEnable
    Case "MASTER DEVICE"
        mnu_mst_device.Enabled = blnEnable
    Case "MASTER ENROLL"
        mnu_mst_enroll.Enabled = blnEnable
    Case "MASTER REPORT"
        mnu_rpt_master.Enabled = blnEnable
            
    Case "SUMMARY LEAVE REPORT"
        mnu_rpt_summary_leave.Enabled = blnEnable
    Case "SUMMARY LEAVE"
        mnu_trans_summary_leave.Enabled = blnEnable
    Case "EMPLOYEE PERFORMANCE"
        mnu_trans_employee_performance.Enabled = blnEnable
                
End Select
End Sub

