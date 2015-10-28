VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_rpt_man_power 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MAN POWER"
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
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3360
      Picture         =   "frm_rpt_man_power.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1890
      Width           =   975
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2340
      Picture         =   "frm_rpt_man_power.frx":058A
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
      Begin MSComCtl2.DTPicker DTPicker_periode 
         Height          =   300
         Left            =   1200
         TabIndex        =   3
         Top             =   660
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   94109699
         UpDown          =   -1  'True
         CurrentDate     =   39278
      End
   End
End
Attribute VB_Name = "frm_rpt_man_power"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim tgl As String
Dim a As New frm_rpt

tgl = Format(DTPicker_Periode.Value, "yyyy-MM-01")

'    If SSTab1.Tab = 1 Then
'        If cbo_periode_to.Text = "To" Then
'            tgl1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
'            tgl2 = Format(DTPicker_periode_to.Value, "yyyy-MM-dd")
'        Else
'            tgl1 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
'            tgl2 = Format(DTPicker_periode_from.Value, "yyyy-MM-dd")
'        End If
'        kdkaryawan = txt_periode_employee_code.Text
'
'    ElseIf SSTab1.Tab = 0 Then
'        If check_valid_yearly Then
'            tgl1 = Format(DTPicker_yearly.Value, "yyyy") & "-01-01"
'            tgl2 = Format(DTPicker_periode_from.Value, "yyyy") & "-12-31"
'            kdkaryawan = txt_yearly_employee_code.Text
'
'        Else
'            Exit Sub
'        End If
'
'    End If
    
    strsql = "call spr_man_power('" & tgl & "')"
    str_file = "\report\rpt_man_power.rpt"
    
    Call a.Show
    a.Caption = "REPORT MAN POWER"
    'str_param_periode = "DAY : (" & Format(DTPicker_daily.Value, "yyyy-MM-dd") & ")"
     
    Call a.rpt_view(strsql, str_file, tgl)
End Sub

Private Sub Form_Load()
DTPicker_Periode = Now

Call load_data_user_access(Me)
End Sub
