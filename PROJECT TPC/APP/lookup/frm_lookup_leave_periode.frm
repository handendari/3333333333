VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_lookup_leave_periode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GENERATE - SUMMARY LEAVE"
   ClientHeight    =   2775
   ClientLeft      =   -15
   ClientTop       =   405
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frm_lookup_leave_periode.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin MSComCtl2.DTPicker DTPicker_periode 
         Height          =   300
         Left            =   1230
         TabIndex        =   1
         Top             =   630
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   104464387
         CurrentDate     =   39278
      End
   End
   Begin prj_tpc.vbButton cmd_use 
      Height          =   450
      Left            =   1650
      TabIndex        =   2
      Top             =   1950
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
      BTYPE           =   14
      TX              =   "&Ok"
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
      MICON           =   "frm_lookup_leave_periode.frx":058A
      PICN            =   "frm_lookup_leave_periode.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmd_close 
      Height          =   450
      Left            =   3000
      TabIndex        =   3
      Top             =   1950
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
      BTYPE           =   14
      TX              =   "&Cancel"
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
      MICON           =   "frm_lookup_leave_periode.frx":1638
      PICN            =   "frm_lookup_leave_periode.frx":1654
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
