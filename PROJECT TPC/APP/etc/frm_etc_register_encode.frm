VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_etc_license_encoder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LICENSE KEY ENCODER"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frm_etc_register_encode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_save 
      Caption         =   "&Save"
      Height          =   645
      Left            =   3960
      Picture         =   "frm_etc_register_encode.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_load 
      Caption         =   "&Local Finger"
      Height          =   645
      Left            =   1800
      Picture         =   "frm_etc_register_encode.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Height          =   315
      Left            =   480
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmd_request_key 
      Caption         =   "Req. Key"
      Height          =   645
      Left            =   480
      Picture         =   "frm_etc_register_encode.frx":0B20
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   6135
      Begin VB.TextBox txt_product_key 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   10
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txt_license_key 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txt_finger_print 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   5
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PRODUCT KEY"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "LICENSE KEY"
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FINGER PRINT"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   1140
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   645
      Left            =   5640
      Picture         =   "frm_etc_register_encode.frx":10AA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmd_register 
      Caption         =   "&Generate"
      Height          =   645
      Left            =   2880
      Picture         =   "frm_etc_register_encode.frx":1634
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   480
      Picture         =   "frm_etc_register_encode.frx":1BBE
      Stretch         =   -1  'True
      Top             =   360
      Width           =   6120
   End
End
Attribute VB_Name = "frm_etc_license_encoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
Unload Me
End Sub



Private Sub cmd_load_Click()
txt_finger_print = get_finger_print
End Sub

Private Sub cmd_register_Click()
txt_license_key = RC4DeCryptASC(txt_finger_print, txt_product_key)
End Sub


Private Sub cmd_save_Click()
Dim i As Boolean

If Trim(txt_license_key) = "" Then Exit Sub
If write_file(App.Path & "\l_" & Format(Now, "yymmddhhnnss") & ".txt") Then
    MsgBox "License key file was successfully created!", vbInformation, "CREATE FILE SUCCESS"
Else
    MsgBox "Error create license key file!", vbCritical, "CREATE FILE ERROR"
End If
End Sub

Private Function write_file(ByVal spath As String) As Boolean
On Error GoTo err
Dim str_data, str1, str2 As String, i As Integer

Open spath For Output As #1
Print #1, txt_license_key.Text
Close #1

write_file = True

Exit Function
err:
'    MsgBox err.Description
    write_file = False
End Function



Private Sub CmdExit_Click()
End
End Sub

Private Sub Form_Load()
txt_product_key = pEncryptionPassword
End Sub
