VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_progess1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processing..."
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
End
Attribute VB_Name = "frm_progess1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bln_act As Boolean


Public Sub show_progrees(ByVal a As Long)
On Error Resume Next
Dim lng_rec_count As Long

Dim lng_jml_btg As Long
lng_rec_count = frm_trans_exp_enroll.Adodc1.Recordset.RecordCount
lng_jml_btg = (a / lng_rec_count) * 100
ProgressBar1.Value = lng_jml_btg
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
bln_act = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Call frm_master_user.cmd_refresh_Click
'Call frm_master_user.cmdConnect_source_Click

Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
If bln_act = True Then
    bln_act = False
Else
    frm_trans_exp_enroll.Adodc1.Recordset.MoveNext
End If

If frm_trans_exp_enroll.Adodc1.Recordset.EOF Then
    Unload Me
    Exit Sub
End If

If frm_trans_exp_enroll.Adodc1.Recordset.Fields("enrollnumber").Value >= lng_start And _
    frm_trans_exp_enroll.Adodc1.Recordset.Fields("enrollnumber").Value <= lng_end Then _
    Call frm_trans_exp_enroll.export_to_device
Call show_progrees(frm_trans_exp_enroll.Adodc1.Recordset.AbsolutePosition)
End Sub
