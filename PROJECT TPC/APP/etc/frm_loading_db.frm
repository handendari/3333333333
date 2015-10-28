VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_loading_db 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm_loading_db.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3680
      _ExtentX        =   6482
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Max             =   12
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Loading database for the first time..."
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   2520
   End
End
Attribute VB_Name = "frm_loading_db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str_license As String



