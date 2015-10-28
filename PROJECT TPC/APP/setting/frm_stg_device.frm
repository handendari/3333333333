VERSION 5.00
Begin VB.Form frm_stg_device 
   Caption         =   "SETTING"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_save 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4560
      Picture         =   "frm_stg_device.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   5760
      Picture         =   "frm_stg_device.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "SETTING ENROLL NUMBER (ACTIVE)"
      Height          =   1815
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   6135
      Begin VB.TextBox txt_en_end 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txt_en_start 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "End"
         Height          =   195
         Left            =   2040
         TabIndex        =   5
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start"
         Height          =   195
         Left            =   2040
         TabIndex        =   4
         Top             =   480
         Width           =   330
      End
   End
End
Attribute VB_Name = "frm_stg_device"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset


Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_save_Click()
With rs
    .Fields("en_number_start").Value = CLng(txt_en_start)
    .Fields("en_number_end").Value = CLng(txt_en_end)
    .Update
End With

lng_start = CLng(txt_en_start)
lng_end = CLng(txt_en_end)
End Sub

Private Sub Form_Load()
Call load_data
End Sub

Private Sub load_data()
If rs.State = 1 Then rs.Close
rs.Open "select * from s_enroll_number where no_urut=1", CnG, adOpenKeyset, adLockOptimistic

If rs.RecordCount > 0 Then
    txt_en_start = rs.Fields("en_number_start").Value
    txt_en_end = rs.Fields("en_number_end").Value
Else
    txt_en_start = 1
    txt_en_end = 100
End If
End Sub
