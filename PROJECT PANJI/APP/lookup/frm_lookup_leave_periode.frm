VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_lookup_leave_periode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GENERATE - REKAP CUTI"
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
      Caption         =   "&Tutup"
      Height          =   645
      Left            =   3360
      Picture         =   "frm_lookup_leave_periode.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmd_use 
      Caption         =   "&OK"
      Height          =   645
      Left            =   2280
      Picture         =   "frm_lookup_leave_periode.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin MSComCtl2.DTPicker DTPicker_periode 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   202440707
         CurrentDate     =   39278
      End
   End
End
Attribute VB_Name = "frm_lookup_leave_periode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_use_Click()
Call frm_trans_leave.generate_summary_leave(DTPicker_periode)
Call cmd_close_Click
End Sub

Private Sub Form_Load()
DTPicker_periode = Now
End Sub
