VERSION 5.00
Begin VB.Form frm_etc_about 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ABOUT"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6990
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5580
      TabIndex        =   2
      Top             =   3360
      Width           =   1030
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   360
      Picture         =   "frmAbout.frx":000C
      Stretch         =   -1  'True
      Top             =   840
      Width           =   6240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   " :: ATTENDANCE && PAYROLL 2.0 :: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1890
      TabIndex        =   9
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   ":  info@solusisentraldata.com"
      Height          =   195
      Left            =   2760
      TabIndex        =   8
      Top             =   2640
      Width           =   2070
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   ":  KTC COAL MINING && ENERGY, PT."
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   2400
      Width           =   2745
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   ":  ©2009 - 2012"
      Height          =   195
      Left            =   2760
      TabIndex        =   6
      Top             =   2880
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   ":  SSD Software Team"
      Height          =   195
      Left            =   2760
      TabIndex        =   5
      Top             =   2160
      Width           =   1590
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Licensed to"
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   2400
      Width           =   825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Email info"
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   675
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   6600
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Copyright"
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   2880
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Created by"
      Height          =   195
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   765
   End
End
Attribute VB_Name = "frm_etc_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

