VERSION 5.00
Begin VB.Form frm_lookup_sort_employee 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SORT EMPLOYEE"
   ClientHeight    =   3510
   ClientLeft      =   -15
   ClientTop       =   345
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_close 
      Caption         =   "&Close"
      Height          =   645
      Left            =   3450
      Picture         =   "frm_lookup_sort_employee.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmd_use 
      Caption         =   "&OK"
      Height          =   645
      Left            =   2370
      Picture         =   "frm_lookup_sort_employee.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sort Employee By"
      Height          =   2475
      Left            =   360
      TabIndex        =   0
      Top             =   150
      Width           =   4065
      Begin VB.OptionButton Option1 
         Caption         =   "Employee Title"
         Height          =   315
         Index           =   4
         Left            =   270
         TabIndex        =   7
         Top             =   1920
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Division"
         Height          =   315
         Index           =   3
         Left            =   270
         TabIndex        =   6
         Top             =   1530
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Department"
         Height          =   315
         Index           =   2
         Left            =   270
         TabIndex        =   5
         Top             =   1140
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Employee Name"
         Height          =   315
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   750
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Caption         =   "No Badge"
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   390
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frm_lookup_sort_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_use_Click()
'Call frm_trans_summary_leave.generate_summary_leave(DTPicker_periode)
'Call cmd_close_Click
Dim str_file As String
Dim a As New frm_rpt
Dim strsql As String, active As String
Dim vorder As String

vorder = IIf(Option1(0), "a.employee_code", IIf(Option1(1), "a.employee_name", _
    IIf(Option1(2), "c.department_name", IIf(Option1(3), "d.division_name", "e.title_name"))))

str_file = "\report\rpt_master_employee.rpt"

active = IIf(frm_mst_employee.optActive.Value, "<> 0", "= 0")
access = IIf(LOGIN_LEVEL = 100, "", "AND (a.managerial_access = 0 OR a.managerial_access IS NULL)")

strsql = "SELECT a.company_code,b.company_name,a.department_code,c.department_name," _
            & "a.employee_code,a.employee_name,a.title_code,e.title_name,a.marital_status," _
            & "a.sex,a.start_working,a.bank_account,a.branches,a.flag_active,a.division_code," _
            & "d.division_name,a.number_of_children,a.religion,a.npwp,a.phone_number," _
            & "a.place_of_birth,a.date_of_birth,a.address " _
            & "FROM m_employee a LEFT JOIN m_company b ON a.company_code = b.company_code " _
            & "LEFT JOIN m_department c ON a.department_code = c.department_code AND c.company_code = b.company_code " _
            & "LEFT JOIN m_division d ON a.division_code = d.division_code AND d.department_code = c.department_code AND d.company_code = b.company_code " _
            & "LEFT JOIN m_title e ON a.title_code = e.title_code " _
            & "where a.company_code = '" _
            & frm_mst_employee.TDBCombo_company.Columns("company_code").Value & "' " _
            & "AND a.flag_active " & active & " " & access & " " _
            & "ORDER BY " & vorder & " ASC"
            
Call a.Show

a.Caption = "REPORT MASTER EMPLOYEE"
Call a.rpt_view(strsql, str_file, Now)

Call cmd_close_Click
End Sub

Private Sub Form_Load()
Option1(0).Value = True
End Sub
