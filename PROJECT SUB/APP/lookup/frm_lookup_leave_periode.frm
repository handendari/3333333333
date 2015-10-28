VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_lookup_leave_periode 
   Caption         =   "GENERATE - SUMMARY LEAVE"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_close 
      Caption         =   "&Close"
      Height          =   645
      Left            =   3360
      Picture         =   "frm_lookup_leave_periode.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmd_use 
      Caption         =   "&Use"
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
         Format          =   97976323
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
Call frm_trans_leave_periode.generate_summary_leave(DTPicker_periode)
Call cmd_close_Click
End Sub

Private Sub Form_Load()
DTPicker_periode = Now
End Sub
