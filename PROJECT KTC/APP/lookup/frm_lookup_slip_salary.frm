VERSION 5.00
Begin VB.Form frm_lookup_slip_salary 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CHOOSE - SLIP SALARY"
   ClientHeight    =   2775
   ClientLeft      =   -15
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_close 
      Caption         =   "&Close"
      Height          =   645
      Left            =   3360
      Picture         =   "frm_lookup_slip_salary.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmd_use 
      Caption         =   "&Use"
      Height          =   645
      Left            =   2280
      Picture         =   "frm_lookup_slip_salary.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1890
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frm_lookup_slip_salary.frx":0B14
         Left            =   1380
         List            =   "frm_lookup_slip_salary.frx":0B1E
         TabIndex        =   3
         Text            =   "Slip Salary Standard"
         Top             =   660
         Width           =   2235
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Choose"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   690
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frm_lookup_slip_salary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_close_Click()
Unload Me
End Sub

'Private Sub cmd_use_Click()
'If Combo1.ListIndex = 0 Then
'    Call frm_rpt_summary_salary.cmd_print_slip_standard
'Else
'    Call frm_rpt_summary_salary.cmd_print_slip_detail
'End If
'
'Call cmd_close_Click
'End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
End Sub
