VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_attendance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup / Restore And Download"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9075
   Icon            =   "frm_attendance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5715
      Left            =   120
      TabIndex        =   1
      Top             =   210
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   10081
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Data Fingerprint"
      TabPicture(0)   =   "frm_attendance.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(1)=   "fra_button_control(1)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Download Log"
      TabPicture(1)   =   "frm_attendance.frx":05A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fra_button_control(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdSakti"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ProgressBar2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Transfer Enroll Data"
      TabPicture(2)   =   "frm_attendance.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).Control(1)=   "fra_button_control(2)"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Transfer Enroll Data"
         Height          =   1995
         Index           =   2
         Left            =   -73290
         TabIndex        =   29
         Top             =   1230
         Width           =   5355
         Begin VB.TextBox txt_employee_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2850
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   37
            Top             =   660
            Width           =   2085
         End
         Begin VB.TextBox txt_device_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   2
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   31
            Top             =   1050
            Width           =   2445
         End
         Begin VB.TextBox txt_port 
            Height          =   285
            Index           =   2
            Left            =   4650
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   1050
            Visible         =   0   'False
            Width           =   345
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_device 
            Height          =   375
            Index           =   2
            Left            =   1140
            OleObjectBlob   =   "frm_attendance.frx":05DE
            TabIndex        =   32
            Top             =   1050
            Width           =   1335
         End
         Begin prj_absensi.vbButton cmdBrowse 
            Height          =   315
            Left            =   2460
            TabIndex        =   36
            Top             =   660
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BTYPE           =   14
            TX              =   "..."
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            MICON           =   "frm_attendance.frx":29D2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txt_employee_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1140
            TabIndex        =   34
            Top             =   660
            Width           =   1305
         End
         Begin VB.Label lblTransfer 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   40
            Top             =   1590
            Width           =   5145
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE*"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   35
            Top             =   1620
            Width           =   4425
         End
         Begin VB.Label lblName 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   33
            Top             =   1080
            Width           =   885
         End
      End
      Begin VB.Frame fra_button_control 
         Height          =   1095
         Index           =   2
         Left            =   -73290
         TabIndex        =   27
         Top             =   3210
         Width           =   5355
         Begin prj_absensi.vbButton cmdTransfer 
            Height          =   750
            Left            =   4110
            TabIndex        =   28
            Top             =   210
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1323
            BTYPE           =   14
            TX              =   "&OK"
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
            MICON           =   "frm_attendance.frx":29EE
            PICN            =   "frm_attendance.frx":2A0A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   135
         Left            =   1950
         TabIndex        =   26
         Top             =   4410
         Visible         =   0   'False
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin prj_absensi.vbButton cmdSakti 
         Height          =   315
         Left            =   5610
         TabIndex        =   25
         Top             =   4050
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   "Proses"
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
         MICON           =   "frm_attendance.frx":3A9C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame1 
         Caption         =   "Download Log"
         Height          =   1995
         Index           =   1
         Left            =   1980
         TabIndex        =   16
         Top             =   990
         Width           =   4695
         Begin VB.Frame fra_progress 
            Height          =   1155
            Left            =   390
            TabIndex        =   21
            Top             =   390
            Visible         =   0   'False
            Width           =   4035
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   135
               Left            =   90
               TabIndex        =   24
               Top             =   930
               Visible         =   0   'False
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   238
               _Version        =   393216
               Appearance      =   0
               Scrolling       =   1
            End
            Begin VB.Label lbl_progress 
               AutoSize        =   -1  'True
               Caption         =   "Progress Status..."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   240
               TabIndex        =   22
               Top             =   450
               Width           =   1845
            End
         End
         Begin VB.TextBox txt_device_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   1
            Left            =   2130
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   18
            Top             =   870
            Width           =   2445
         End
         Begin VB.TextBox txt_port 
            Height          =   285
            Index           =   1
            Left            =   4320
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   870
            Visible         =   0   'False
            Width           =   345
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_device 
            Height          =   375
            Index           =   1
            Left            =   780
            OleObjectBlob   =   "frm_attendance.frx":3AB8
            TabIndex        =   19
            Top             =   870
            Width           =   1335
         End
         Begin VB.Label lblName 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DEVICE :"
            Height          =   255
            Index           =   1
            Left            =   -150
            TabIndex        =   20
            Top             =   900
            Width           =   885
         End
      End
      Begin VB.Frame fra_button_control 
         Height          =   1095
         Index           =   1
         Left            =   -72930
         TabIndex        =   14
         Top             =   3270
         Width           =   4695
         Begin prj_absensi.vbButton cmdOK 
            Height          =   750
            Left            =   3540
            TabIndex        =   15
            Top             =   210
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1323
            BTYPE           =   14
            TX              =   "&Backup"
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
            MICON           =   "frm_attendance.frx":5EAC
            PICN            =   "frm_attendance.frx":5EC8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame fra_button_control 
         Height          =   1095
         Index           =   0
         Left            =   1980
         TabIndex        =   10
         Top             =   2940
         Width           =   4695
         Begin prj_absensi.vbButton cmd_download 
            Height          =   705
            Left            =   1620
            TabIndex        =   11
            Top             =   210
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Download"
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
            MICON           =   "frm_attendance.frx":6F5A
            PICN            =   "frm_attendance.frx":6F76
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_absensi.vbButton cmd_delete_log 
            Height          =   705
            Left            =   2610
            TabIndex        =   12
            Top             =   210
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Delete"
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
            MICON           =   "frm_attendance.frx":8008
            PICN            =   "frm_attendance.frx":8024
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_absensi.vbButton cmd_reproses 
            Height          =   705
            Left            =   3600
            TabIndex        =   13
            Top             =   210
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Reprocess"
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
            MICON           =   "frm_attendance.frx":90B6
            PICN            =   "frm_attendance.frx":90D2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Backup and Restore"
         Height          =   1995
         Index           =   0
         Left            =   -72930
         TabIndex        =   2
         Top             =   1290
         Width           =   4695
         Begin VB.OptionButton optType 
            Caption         =   "Restore"
            Height          =   195
            Index           =   1
            Left            =   2490
            TabIndex        =   6
            Top             =   390
            Width           =   915
         End
         Begin VB.OptionButton optType 
            Caption         =   "Backup"
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   5
            Top             =   390
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.TextBox txt_port 
            Height          =   285
            Index           =   0
            Left            =   4290
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   870
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txt_device_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   0
            Left            =   2130
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            Top             =   870
            Width           =   2445
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_device 
            Height          =   375
            Index           =   0
            Left            =   780
            OleObjectBlob   =   "frm_attendance.frx":A164
            TabIndex        =   7
            Top             =   870
            Width           =   1335
         End
         Begin VB.Label lblName 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DARI :"
            Height          =   255
            Index           =   0
            Left            =   -150
            TabIndex        =   9
            Top             =   900
            Width           =   885
         End
         Begin VB.Label lblCaption 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   8
            Top             =   1620
            Width           =   4425
         End
      End
   End
   Begin prj_absensi.LynxGrid LynxGrid2 
      Height          =   3255
      Left            =   2970
      TabIndex        =   39
      Top             =   2430
      Visible         =   0   'False
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   5741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorSel    =   12937777
      ForeColorSel    =   16777215
      CustomColorFrom =   16572875
      CustomColorTo   =   14722429
      GridColor       =   16367254
      FocusRectColor  =   9895934
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      AllowColumnResizing=   -1  'True
   End
   Begin VB.Timer timer_get_log_data 
      Interval        =   60000
      Left            =   8280
      Top             =   5400
   End
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   345
      Left            =   150
      OleObjectBlob   =   "frm_attendance.frx":C558
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   405
   End
   Begin prj_absensi.vbButton cmdExit 
      Height          =   705
      Left            =   7650
      TabIndex        =   23
      Top             =   5970
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1244
      BTYPE           =   14
      TX              =   "&Exit"
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
      MICON           =   "frm_attendance.frx":C57C
      PICN            =   "frm_attendance.frx":C598
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frm_attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE       As Long = &H100000
Private Const INFINITE          As Long = &HFFFF

Dim rsDevice As New ADODB.Recordset
Dim cn_fp As Boolean

Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim lngErrorCode As Long
Dim bConnected As Boolean
Dim connFP As New ADODB.Connection
Dim recFP As New ADODB.Recordset
Dim OnFingerTime As Long
Dim VerifyTime As Long
Dim vMachineNumber
Public vAutoDeleteLog As Integer
Dim vMode As Integer

Dim public_int_caller As Integer
Dim rs_bound As New ADODB.Recordset

Dim vGroupCode As String
Dim vShiftCode As String

Private Function connect() As Boolean
    If cn_fp Then
        CZKEM1.EnableDevice vMachineNumber, True
        CZKEM1.disconnect
    End If
    
    cn_fp = CZKEM1.Connect_Net(FG_IP_ADDRESS, CLng(FG_PORT_NUMBER))
    
    If cn_fp Then
        CZKEM1.EnableDevice vMachineNumber, False
        connect = True
    Else
        connect = False
        Exit Function
    End If
End Function

Private Function disconnect() As Boolean
    If cn_fp Then
        CZKEM1.EnableDevice vMachineNumber, True
        CZKEM1.disconnect
    End If
End Function

Private Function clear_all_log() As Boolean
Dim vRet As Boolean
Dim vErrorCode As Long

    vRet = CZKEM1.ClearGLog(vMachineNumber)
    If vRet Then
        clear_all_log = True
    Else
        clear_all_log = False
    End If
End Function

Private Sub cmd_download_Click()
Dim rscari As New ADODB.Recordset
Dim rsMesin As New ADODB.Recordset

Dim x
Dim alg_ver As String
Dim iMachineNumber As Long

    FG_IP_ADDRESS = TDBCombo_device(1).Columns("ip_address").Value
    FG_PORT_NUMBER = TDBCombo_device(1).Columns("port_number").Value
    
    If rsMesin.State Then rsMesin.Close
    SQL = "SELECT auto_del_log FROM m_device WHERE ip_address = '" & FG_IP_ADDRESS & "'"
    rsMesin.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsMesin.RecordCount > 0 Then
        vAutoDeleteLog = Val("" & rsMesin.Fields("auto_del_log").Value)
    End If
    rsMesin.Close
    
    Call connect
    
    x = CZKEM1.EnableDevice(iMachineNumber, False)
        
    'Only TFT screen devices with firmware version Ver 6.60 version later support function "GetUserTmpExStr" and "GetUserTmpEx".
    'While you are using 9.0 fingerprint arithmetic and your device's firmware version is under ver6.60,you should use the functions "SSR_GetUserTmp" or
    '"SSR_GetUserTmpStr" instead of "GetUserTmpExStr" or "GetUserTmpEx" in order to download the fingerprint templates.
    
    x = CZKEM1.GetSysOption(CLng(vMachineNumber), "~ZKFPVersion", alg_ver)
    If (alg_ver = "9" Or alg_ver = "") And (CZKEM1.IsTFTMachine(vMachineNumber) = False) Then 'Alg 9
        fra_progress.Visible = True: fra_button_control(0).Enabled = False
        lbl_progress.Caption = "Read Data Fingerprint..."
        Call disconnect
        Call download_action_typeA
    Else
        fra_progress.Visible = True: fra_button_control(0).Enabled = False
        lbl_progress.Caption = "Read Data Fingerprint..."
        Call disconnect
        Call download_action_typeB
    End If
    
'    If vAutoDeleteLog <> 0 Then
'        Call delete_log_action
'    End If
End Sub

Private Sub cmd_delete_log_Click()
    Call delete_log_action
End Sub

Private Sub cmd_reproses_Click()
    frm_reproccess_download.Show
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim dwEnrollNmber As Long, name As String, passWord As String, privilege As Long, Enabled As Boolean
Dim dwEnrollNmberTFT As String

Dim iEnrollNumber As Long
Dim iEnrollNumberTFT As String
Dim iBackupNumber As Long
Dim iMachineNumber As Long
Dim sTmpData As String
Dim Flag As Long, tmp_len As Long

Dim iUsername As String, iPassword As String

Dim sEnrollData As String
Dim iLength As Long
Dim alg_ver As String
Dim tmpLength As Long

Dim vFlagSuccess As Integer
Dim vFlagFailed As Integer

Dim x
    
On Error Resume Next
    FG_IP_ADDRESS = TDBCombo_device(0).Columns("ip_address").Value
    FG_PORT_NUMBER = TDBCombo_device(0).Columns("port_number").Value
    
    vFlagSuccess = 0
    vFlagFailed = 0
    
    Me.MousePointer = vbHourglass
    If connect Then
        If optType(0) Then
                    
            tmp_len = 1024 * 20 '//Fixed Value
            iBackupNumber = 0
            x = CZKEM1.EnableDevice(iMachineNumber, False)
        
            'Only TFT screen devices with firmware version Ver 6.60 version later support function "GetUserTmpExStr" and "GetUserTmpEx".
            'While you are using 9.0 fingerprint arithmetic and your device's firmware version is under ver6.60,you should use the functions "SSR_GetUserTmp" or
            '"SSR_GetUserTmpStr" instead of "GetUserTmpExStr" or "GetUserTmpEx" in order to download the fingerprint templates.
            
            x = CZKEM1.GetSysOption(CLng(vMachineNumber), "~ZKFPVersion", alg_ver)
            If CZKEM1.ReadAllUserID(vMachineNumber) Then
                If alg_ver = "9" And (CZKEM1.IsTFTMachine(vMachineNumber) = False) Then 'Alg 9
                    While CZKEM1.GetAllUserInfo(vMachineNumber, dwEnrollNmber, name, passWord, privilege, Enabled)
                        For iBackupNumber = 0 To 9
                            If CZKEM1.GetUserTmpExStr(CLng(vMachineNumber), CStr(dwEnrollNmber), CLng(iBackupNumber), 1, sTmpData, CLng(tmp_len)) Then
                                'Add to DB
                                recFP.Open "select * from fptable", CnG, adOpenKeyset, adLockOptimistic
                                recFP.AddNew
                                recFP.Fields("EMachineNumber") = vMachineNumber
                                recFP.Fields("EnrollNumber") = dwEnrollNmber
                                recFP.Fields("username") = name
                                recFP.Fields("Password") = passWord
                                recFP.Fields("Privilege") = privilege
                                recFP.Fields("FingerNumber") = iBackupNumber
                                recFP.Fields("Template") = sTmpData
                                recFP.Update
                                recFP.Close
                                
                                vFlagSuccess = vFlagSuccess + 1
                                
                                lblCaption.Caption = "Backup Data " & vFlagSuccess
                            End If
                        Next iBackupNumber
                        DoEvents
                        
                    Wend
                                    
                    MsgBox vFlagSuccess & " Backup data successfully...", vbInformation, headerMSG
                Else
                    While CZKEM1.SSR_GetAllUserInfo(CLng(vMachineNumber), dwEnrollNmberTFT, name, passWord, privilege, Enabled)
                        For iBackupNumber = 0 To 9
                            If CZKEM1.SSR_GetUserTmpStr(CLng(vMachineNumber), dwEnrollNmberTFT, iBackupNumber, sTmpData, CLng(tmp_len)) Then
                                'Add to DB
                                recFP.Open "select * from fptable", CnG, adOpenKeyset, adLockOptimistic
                                recFP.AddNew
                                recFP.Fields("EMachineNumber") = vMachineNumber
                                recFP.Fields("EnrollNumber") = dwEnrollNmberTFT
                                recFP.Fields("username") = name
                                recFP.Fields("Password") = passWord
                                recFP.Fields("Privilege") = privilege
                                recFP.Fields("FingerNumber") = iBackupNumber
                                recFP.Fields("Template") = sTmpData
                                recFP.Update
                                recFP.Close
                                
                                vFlagSuccess = vFlagSuccess + 1
                                lblCaption.Caption = "Backup Data " & vFlagSuccess
                            End If
                        Next iBackupNumber
                        DoEvents
                    Wend
                    
                    MsgBox vFlagSuccess & " Backup data successfully...", vbInformation, headerMSG
                End If
            End If
            x = CZKEM1.EnableDevice(iMachineNumber, True)
        Else
'            If CZKEM1.IsTFTMachine(vMachineNumber) = False Then '//bw
'                x = CZKEM1.SetUserInfo(CLng(vMachineNumber), CLng(iEnrollNumber), "Staff Name", "", 0, True)
'            Else
'                x = CZKEM1.SSR_SetUserInfo(vMachineNumber, iEnrollNumber, "Staff Name", "", 0, True)
'            End If
            
            'Read Template and index finger From DB
            SQL = "select * from  fptable"
            recFP.Open SQL, CnG, adOpenKeyset, adLockOptimistic
            
            If recFP.RecordCount > 0 Then
                recFP.MoveFirst
                While Not recFP.EOF
                    sEnrollData = Trim(recFP("Template"))
                    iEnrollNumber = recFP("EnrollNumber")
                    iEnrollNumberTFT = recFP("EnrollNumber")
                    iBackupNumber = recFP("FingerNumber")
                    iUsername = recFP("username")
                    iPassword = recFP("Password")
                    
                    If CZKEM1.IsTFTMachine(vMachineNumber) = False Then '//bw
                        x = CZKEM1.SetUserInfo(CLng(vMachineNumber), iEnrollNumber, iUsername, iPassword, 0, True)
                    Else
                        x = CZKEM1.SSR_SetUserInfo(vMachineNumber, iEnrollNumberTFT, iUsername, iPassword, 0, True)
                    End If

                    x = CZKEM1.GetSysOption(CLng(vMachineNumber), "~ZKFPVersion", alg_ver)
            
                    If CZKEM1.ReadAllUserID(vMachineNumber) Then
                        If alg_ver = "9" And (CZKEM1.IsTFTMachine(vMachineNumber) = False) Then 'Alg 9
                            If CZKEM1.SetUserTmpStr(vMachineNumber, iEnrollNumber, iBackupNumber, sEnrollData) Then
                                vFlagSuccess = vFlagSuccess + 1
                            Else
                                vFlagFailed = vFlagFailed + 1
                            End If
                        Else
                            If CZKEM1.SetUserTmpExStr(vMachineNumber, iEnrollNumberTFT, iBackupNumber, 1, sEnrollData) Then
                                vFlagSuccess = vFlagSuccess + 1
                            Else
                                vFlagFailed = vFlagFailed + 1
                            End If
                        End If
                    End If
                    
                    lblCaption.Caption = "Restore Data " & vFlagSuccess
                recFP.MoveNext
                DoEvents
                Wend
                
                MsgBox vFlagSuccess & " Restore data successfully......" & Chr(13) & _
                        vFlagFailed & " Restore data failed......", vbInformation, headerMSG
                        
                lblCaption.Caption = ""
            End If
            recFP.Close
        End If
        Call disconnect
    Else
        MsgBox "Error to connecting fingerprint...", vbExclamation, headerMSG
        Me.MousePointer = vbNormal
        Exit Sub
    End If
    Me.MousePointer = vbNormal
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdSakti_Click()
Dim vParameter As String
Dim vTglAwal As Date
Dim vTglAkhir As Date

Dim vFlagType As String
Dim vDescription As String

'On Error GoTo Err
    
    CnG.BeginTrans
    
    '+++++++++++++++++++++++++++++ LEAVE +++++++++++++++++++++++++++++++++++++++++
    If rs.State Then rs.Close
    SQL = "SELECT * FROM t_leave WHERE flag_date_to = 0"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    ProgressBar2.Visible = True
    ProgressBar2.Value = 0
        
    If rs.RecordCount > 0 Then
        ProgressBar2.Max = rs.RecordCount
        
        rs.MoveFirst
        While Not rs.EOF
            DoEvents
            ProgressBar2.Value = ProgressBar2.Value + 1
                
            vTglAwal = Format(rs!leave_date_from.Value, "yyyy-MM-dd")
            vTglAkhir = Format(rs!leave_date_from.Value, "yyyy-MM-dd")
            
            vTglAwal = DateValue(vTglAwal)
            vTglAkhir = DateValue(vTglAkhir)
                        
            While vTglAwal <= vTglAkhir
                SQL = "DELETE FROM h_attendance WHERE employee_code = '" & rs!employee_code & "' " & _
                        "AND DATE(att_date) = '" & Format(vTglAwal, "yyyy-MM-dd") & "'"
                CnG.Execute SQL
                
                If rscari.State Then rscari.Close
                SQL = "SELECT a.group_code,b.shift_code " & _
                      "FROM td_shift a JOIN tm_shift b ON a.shift_number = b.shift_number AND a.group_code = b.group_code " & _
                      "WHERE DATE(b.start_date) <= DATE('" & Format(vTglAwal, "yyyy-MM-dd") & "') AND a.employee_code = '" & rs!employee_code & "' " & _
                      "ORDER BY b.start_date DESC LIMIT 1"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vGroupCode = rscari!group_code
                    vShiftCode = rscari!shift_code
                Else
                    vGroupCode = "ST"
                    vShiftCode = "ST01"
                End If
                rscari.Close
                
                SQL = "INSERT INTO h_attendance (employee_code, att_date, shift_number, group_code, shift_code, status, absent_status, description, flag_manual, entry_date) " & _
                      "VALUES (" & _
                        "'" & rs!employee_code & "','" & Format(vTglAwal, "yyyy-MM-dd") & "'," & _
                        "1,'" & vGroupCode & "','" & vShiftCode & "'," & _
                        "CASE WHEN '" & rs!flag_type & "' = 0 THEN 'L' " & _
                            "WHEN '" & rs!flag_type & "' = 1 THEN 'HC' ELSE 'UL' END," & _
                        "CASE WHEN '" & rs!flag_type & "' = 0 THEN 3 " & _
                            "WHEN '" & rs!flag_type & "' = 1 THEN 31 ELSE 32 END," & _
                        "'" & rs!Description & "',1,now())"
                CnG.Execute SQL
                
'                SQL = "UPDATE h_attendance SET STATUS = CASE WHEN '" & rs!flag_type & "' = 0 THEN 'L' " & _
'                        "WHEN '" & rs!flag_type & "' = 1 THEN 'HC' ELSE 'UL' END," & _
'                        "description = '" & rs!Description & "' " & _
'                      "WHERE employee_code = '" & rs!employee_code & "' " & _
'                        "AND DATE(att_date) = '" & Format(vTglAwal, "yyyy-MM-dd") & "'"
'                CnG.Execute SQL
                
                vTglAwal = vTglAwal + 1
            Wend
            
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    If rs.State Then rs.Close
    SQL = "SELECT * FROM t_leave WHERE flag_date_to = 1"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    ProgressBar2.Visible = True
    ProgressBar2.Value = 0
        
    If rs.RecordCount > 0 Then
        ProgressBar2.Max = rs.RecordCount
        
        rs.MoveFirst
        While Not rs.EOF
            DoEvents
            ProgressBar2.Value = ProgressBar2.Value + 1
                
            vTglAwal = Format(rs!leave_date_from.Value, "yyyy-MM-dd")
            vTglAkhir = Format(rs!leave_date_to.Value, "yyyy-MM-dd")
            
            vTglAwal = DateValue(vTglAwal)
            vTglAkhir = DateValue(vTglAkhir)
                        
            While vTglAwal <= vTglAkhir
                SQL = "DELETE FROM h_attendance WHERE employee_code = '" & rs!employee_code & "' " & _
                        "AND DATE(att_date) = '" & Format(vTglAwal, "yyyy-MM-dd") & "'"
                CnG.Execute SQL
                
                If rscari.State Then rscari.Close
                SQL = "SELECT a.group_code,b.shift_code " & _
                      "FROM td_shift a JOIN tm_shift b ON a.shift_number = b.shift_number AND a.group_code = b.group_code " & _
                      "WHERE DATE(b.start_date) <= DATE('" & Format(vTglAwal, "yyyy-MM-dd") & "') AND a.employee_code = '" & rs!employee_code & "' " & _
                      "ORDER BY b.start_date DESC LIMIT 1"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vGroupCode = rscari!group_code
                    vShiftCode = rscari!shift_code
                Else
                    vGroupCode = "ST"
                    vShiftCode = "ST01"
                End If
                rscari.Close
                
                SQL = "INSERT INTO h_attendance (employee_code, att_date, shift_number, group_code, shift_code, status, absent_status, description, flag_manual, entry_date) " & _
                      "VALUES (" & _
                        "'" & rs!employee_code & "','" & Format(vTglAwal, "yyyy-MM-dd") & "'," & _
                        "1,'" & vGroupCode & "','" & vShiftCode & "'," & _
                        "CASE WHEN '" & rs!flag_type & "' = 0 THEN 'L' " & _
                            "WHEN '" & rs!flag_type & "' = 1 THEN 'HC' ELSE 'UL' END," & _
                        "CASE WHEN '" & rs!flag_type & "' = 0 THEN 3 " & _
                            "WHEN '" & rs!flag_type & "' = 1 THEN 31 ELSE 32 END," & _
                        "'" & rs!Description & "',1,now())"
                CnG.Execute SQL
                
'                SQL = "UPDATE h_attendance SET STATUS = CASE WHEN '" & rs!flag_type & "' = 0 THEN 'L' " & _
'                        "WHEN '" & rs!flag_type & "' = 1 THEN 'HC' ELSE 'UL' END," & _
'                        "description = '" & rs!Description & "' " & _
'                      "WHERE employee_code = '" & rs!employee_code & "' " & _
'                        "AND DATE(att_date) = '" & Format(vTglAwal, "yyyy-MM-dd") & "'"
'                CnG.Execute SQL
                
                vTglAwal = vTglAwal + 1
            Wend
            
            rs.MoveNext
        Wend
    End If
    rs.Close
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '++++++++++++++++++++++ SPESIAL LEAVE +++++++++++++++++++++++++++++++++++++++++++
    If rs.State Then rs.Close
    SQL = "SELECT * FROM t_permission"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
    ProgressBar2.Value = 0
        
    If rs.RecordCount > 0 Then
        ProgressBar2.Max = rs.RecordCount
        
        rs.MoveFirst
        While Not rs.EOF
            DoEvents
            ProgressBar2.Value = ProgressBar2.Value + 1
            
            vFlagType = rs!flag_type
            vTglAwal = Format(rs!date_from.Value, "yyyy-MM-dd")
            vTglAkhir = Format(rs!date_to.Value, "yyyy-MM-dd")
            
            vTglAwal = DateValue(vTglAwal)
            vTglAkhir = DateValue(vTglAkhir)
                        
            While vTglAwal <= vTglAkhir
                SQL = "DELETE FROM h_attendance WHERE employee_code = '" & rs!employee_code & "' " & _
                        "AND date(att_date) = '" & Format(vTglAwal, "yyyy-MM-dd") & "'"
                CnG.Execute SQL
                
                If rscari.State Then rscari.Close
                SQL = "SELECT a.group_code,b.shift_code " & _
                      "FROM td_shift a JOIN tm_shift b ON a.shift_number = b.shift_number AND a.group_code = b.group_code " & _
                      "WHERE DATE(b.start_date) <= DATE('" & Format(vTglAwal, "yyyy-MM-dd") & "') AND a.employee_code = '" & rs!employee_code & "' " & _
                      "ORDER BY b.start_date DESC LIMIT 1"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    vGroupCode = rscari!group_code
                    vShiftCode = rscari!shift_code
                Else
                    vGroupCode = "ST"
                    vShiftCode = "ST01"
                End If
                rscari.Close
                
                
                SQL = "INSERT INTO h_attendance (employee_code, att_date, shift_number, group_code, shift_code, status, absent_status, description, flag_manual, entry_date) " & _
                      "VALUES (" & _
                        "'" & rs!employee_code & "','" & Format(vTglAwal, "yyyy-MM-dd") & "',1,'" & vGroupCode & "','" & vShiftCode & "','SL'," & _
                        "" & IIf(vFlagType = 1, 20, IIf(vFlagType = 2, 21, _
                            IIf(vFlagType = 3, 22, IIf(vFlagType = 4, 23, IIf(vFlagType = 5, 24, IIf(vFlagType = 5, 25, _
                            IIf(vFlagType = 6, 26, IIf(vFlagType = 7, 27, 28)))))))) & "," & _
                        "'" & rs!Description & "',1,now())"
                CnG.Execute SQL
                
                vTglAwal = vTglAwal + 1
                
'                SQL = "UPDATE h_attendance SET STATUS = 'SL'," & _
'                        "absen_status = " & IIf(rs!flag_type = 1, 20, IIf(vFlagType = 2, 21, _
'                            IIf(vFlagType = 3, 22, IIf(vFlagType = 4, 23, IIf(vFlagType = 5, 24, IIf(vFlagType = 5, 25, _
'                            IIf(vFlagType = 6, 26, IIf(vFlagType = 7, 27, 28)))))))) & "," & _
'                        "description = '" & rs!Description & "' " & _
'                      "WHERE employee_code = '" & rs!employee_code & "' " & _
'                        "AND DATE(att_date) = '" & Format(vTglAwal, "yyyy-MM-dd") & "'"
'                CnG.Execute SQL
                
                vTglAwal = vTglAwal + 1
            Wend
            
            rs.MoveNext
        Wend
    End If
    rs.Close
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    CnG.CommitTrans
    
    ProgressBar2.Visible = False
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdTransfer_Click()
Dim dwEnrollNmber As Long, name As String, passWord As String, privilege As Long, Enabled As Boolean
Dim dwEnrollNmberTFT As String

Dim iEnrollNumber As Long
Dim iEnrollNumberTFT As String
Dim iBackupNumber As Long
Dim iMachineNumber As Long
Dim sTmpData As String
Dim Flag As Long, tmp_len As Long

Dim iUsername As String, iPassword As String

Dim sEnrollData As String
Dim iLength As Long
Dim alg_ver As String
Dim tmpLength As Long

Dim vFlagSuccess As Integer
Dim vFlagFailed As Integer

Dim x

    If txt_employee_code.Text = "" Then
        MsgBox "Employee is empty...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    If TDBCombo_device(SSTab1.Tab).Text = "" Then
        MsgBox "Device is empty...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    FG_IP_ADDRESS = TDBCombo_device(2).Columns("ip_address").Value
    FG_PORT_NUMBER = TDBCombo_device(2).Columns("port_number").Value
    
    vFlagSuccess = 0
    vFlagFailed = 0
    
    If rs.State Then rs.Close
    SQL = "SELECT DISTINCT enrollnumber FROM m_enroll_link WHERE employee_code = '" & txt_employee_code.Text & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Me.MousePointer = vbHourglass
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            If rscari.State Then rscari.Close
            SQL = "SELECT a.EnrollNumber, a.FingerNumber, c.employee_nick_name EnrollName, a.Password, CAST(a.FPData AS CHAR) FPData " & _
                  "FROM m_enroll a JOIN m_enroll_link b ON a.enrollnumber = b.enrollnumber " & _
                  "JOIN m_employee c ON b.employee_code = c.employee_code " & _
                  "WHERE a.enrollnumber = '" & rs.Fields(0).Value & "' AND IFNULL(b.employee_code,'') <> ''"
            rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rscari.RecordCount > 0 Then
                rscari.MoveFirst
                While Not rscari.EOF
                    If connect Then
                        sEnrollData = Trim(rscari("FPData"))
                        iEnrollNumber = rscari("EnrollNumber")
                        iEnrollNumberTFT = rscari("EnrollNumber")
                        iBackupNumber = rscari("FingerNumber")
                        iUsername = rscari("EnrollName")
                        iPassword = rscari("Password")
                        
                        If CZKEM1.IsTFTMachine(vMachineNumber) = False Then '//bw
                            x = CZKEM1.SetUserInfo(CLng(vMachineNumber), iEnrollNumber, iUsername, iPassword, 0, True)
                        Else
                            x = CZKEM1.SSR_SetUserInfo(vMachineNumber, iEnrollNumberTFT, iUsername, iPassword, 0, True)
                        End If
    
                        x = CZKEM1.GetSysOption(CLng(vMachineNumber), "~ZKFPVersion", alg_ver)
                
                        If CZKEM1.ReadAllUserID(vMachineNumber) Then
                            If alg_ver = "9" And (CZKEM1.IsTFTMachine(vMachineNumber) = False) Then 'Alg 9
                                If CZKEM1.SetUserTmpStr(vMachineNumber, iEnrollNumber, iBackupNumber, sEnrollData) Then
                                    vFlagSuccess = vFlagSuccess + 1
                                Else
                                    vFlagFailed = vFlagFailed + 1
                                End If
                            Else
                                If CZKEM1.SetUserTmpExStr(vMachineNumber, iEnrollNumberTFT, iBackupNumber, 1, sEnrollData) Then
                                    vFlagSuccess = vFlagSuccess + 1
                                Else
                                    vFlagFailed = vFlagFailed + 1
                                End If
                            End If
                        End If
                        
                        lblTransfer.Caption = "Restore Data " & vFlagSuccess
                    End If
                    rscari.MoveNext
                Wend
            Else
                lblTransfer.Caption = "Enroll Data Not Found"
            End If
            rscari.Close
            
            rs.MoveNext
        Wend
    Else
        lblTransfer.Caption = "Enroll Data Not Found"
    End If
    rs.Close
    
    Call disconnect
    
    Me.MousePointer = vbNormal
End Sub

Private Sub optType_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 81 Then
        End
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call load_data_device
    
    If SSTab1.Tab = 2 Then
        createGridKar
    End If
End Sub

Private Sub TDBCombo_device_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 81 Then
        End
    End If
End Sub

Private Sub txt_device_name_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 81 Then
        End
    End If
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 81 Then
        End
    End If
End Sub

Private Sub cmd_download_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 81 Then
        End
    End If
End Sub

Private Sub cmd_delete_log_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 81 Then
        End
    End If
End Sub

Private Sub Form_Load()
    SSTab1.TabVisible(0) = False
    
    SSTab1.Tab = 1
    
    vMachineNumber = 1
    Call load_data_device
    optType(0) = True
End Sub

Private Sub load_data_device()
    If rsDevice.State = 1 Then rsDevice.Close
    SQL = "select * from m_device order by ip_address"
    rsDevice.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_device(SSTab1.Tab).RowSource = rsDevice
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_attendance = Nothing
End Sub

Private Sub optType_Click(index As Integer)
    If optType(0) Then
        lblName(SSTab1.Tab).Caption = "FROM :"
        cmdOK.Caption = "&Backup"
        cmdOK.Enabled = True
    ElseIf optType(1) Then
        lblName(SSTab1.Tab).Caption = "TO :"
        cmdOK.Caption = "&Restore"
        cmdOK.Enabled = True
    End If
End Sub

Private Sub TDBCombo_device_ItemChange(index As Integer)
    If TDBCombo_device(SSTab1.Tab).ApproxCount > 0 Then
        TDBCombo_device(SSTab1.Tab).Text = TDBCombo_device(SSTab1.Tab).Columns("ip_address").Value
        txt_device_name(SSTab1.Tab).Text = TDBCombo_device(SSTab1.Tab).Columns("name").Value
        txt_port(SSTab1.Tab).Text = TDBCombo_device(SSTab1.Tab).Columns("port").Value
    End If
End Sub

Private Function get_employee_code(ByVal int_number As Long) As String
Dim rs1 As New ADODB.Recordset
Dim str_employee_code As String

    SQL = "SELECT company_code, employee_code  " _
        & "FROM m_enroll_link WHERE ip_address = '" & FG_IP_ADDRESS & _
        "' AND enrollnumber = " & int_number
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open SQL, CnG, adOpenStatic, adLockReadOnly
    If rs1.RecordCount > 0 Then
        str_employee_code = rs1.Fields("employee_code").Value
    Else
        str_employee_code = ""
    End If
    
    get_employee_code = str_employee_code
End Function

'Public Sub download_action_typeA()
'    If rs_bound.State = 1 Then rs_bound.Close
'    rs_bound.Open "select * from h_log_attendance where att_number = -1", CnG, adOpenKeyset, adLockOptimistic
'
'    lbl_progress.Caption = "Downloading log data..."
'    fra_progress.Visible = True: fra_button_control(SSTab1.Tab).Enabled = False
'
'    If Not connect Then
'        If Not public_int_caller = 1 Then MsgBox "Error connecting to source device...", vbCritical, headerMSG
'    Else
'        If download_data_log_typeA Then
'            MsgBox "Downloading is Done !", vbInformation, headerMSG
'        Else
'            MsgBox "Error downloading !", vbCritical, headerMSG
'        End If
'        Call disconnect
'    End If
'
'    fra_progress.Visible = False: fra_button_control(SSTab1.Tab).Enabled = True
'End Sub
'
'Public Sub download_action_typeB()
'    If rs_bound.State = 1 Then rs_bound.Close
'    rs_bound.Open "select * from h_log_attendance where att_number = -1", CnG, adOpenKeyset, adLockOptimistic
'
'    lbl_progress.Caption = "Downloading log data..."
'    fra_progress.Visible = True: fra_button_control(SSTab1.Tab).Enabled = False
'
'    If Not connect Then
'        If Not public_int_caller = 1 Then MsgBox "Error connecting to source device...", vbCritical, headerMSG
'    Else
'        If download_data_log_typeB Then
'            MsgBox "Downloading is Done !", vbInformation, headerMSG
'        Else
'            MsgBox "Error downloading !", vbCritical, headerMSG
'        End If
'        Call disconnect
'    End If
'
'    fra_progress.Visible = False: fra_button_control(SSTab1.Tab).Enabled = True
'End Sub

Public Sub download_action_typeA()
'    lbl_progress.Caption = "Read Data Fingerprint..."
'    fra_progress.Visible = True: fra_button_control.Enabled = False
    
    If Not connect Then
        If Not public_int_caller = 1 Then MsgBox "Error connecting to source device...", vbCritical, headerMSG
    Else
        If download_data_log_typeA Then
            If vMode = 0 Then
                MsgBox "Downloading is Done !", vbInformation, headerMSG
            End If
        Else
            If vMode = 0 Then
                MsgBox "Error downloading !", vbCritical, headerMSG
            End If
        End If
        Call disconnect
    End If
    
    fra_progress.Visible = False: fra_button_control(0).Enabled = True: ProgressBar1.Visible = False
End Sub

Public Sub download_action_typeB()
'    lbl_progress.Caption = "Read Data Fingerprint..."
'    fra_progress.Visible = True: fra_button_control.Enabled = False
    
    If Not connect Then
        If Not public_int_caller = 1 Then MsgBox "Error connecting to source device...", vbCritical, headerMSG
    Else
        If download_data_log_typeB Then
            If vMode = 0 Then
                MsgBox "Downloading is Done !", vbInformation, headerMSG
            End If
        Else
            If vMode = 0 Then
                MsgBox "Error downloading !", vbCritical, headerMSG
            End If
        End If
        Call disconnect
    End If
    
    fra_progress.Visible = False: fra_button_control(0).Enabled = True: ProgressBar1.Visible = False
End Sub

Public Sub delete_log_action()
'    If rs_bound.State = 1 Then rs_bound.Close
'    rs_bound.Open "select * from h_log_attendance where att_number = -1", CnG, adOpenKeyset, adLockOptimistic
    
    lbl_progress.Caption = "Delete log data..."
    fra_progress.Visible = True: fra_button_control(0).Enabled = False
    
    If Not connect Then
        If Not public_int_caller = 1 Then MsgBox "Error connecting to source device...", vbCritical, headerMSG
    Else
        If clear_all_log = False Then
            If Not public_int_caller = 1 Then MsgBox "Error deleting log...", vbCritical, headerMSG
        Else
            If Not public_int_caller = 1 Then MsgBox "Delete log successfully...", vbInformation, headerMSG
        End If
        Call disconnect
    End If
    
    fra_progress.Visible = False: fra_button_control(0).Enabled = True

End Sub

Public Function download_data_log_typeA() As Boolean
On Error Resume Next

Dim dwEnrollNumber As Long
Dim dwVerifyMode As Long
Dim dwInOutMode As Long
Dim timeStr As String
Dim i As Long
Dim lng_year As Long
Dim lng_month As Long
Dim lng_day As Long
Dim lng_hour As Long
Dim lng_minute As Long
Dim lng_second As Long
Dim dw_work As Long
Dim strsql As String
Dim j As Long

Dim rsabsen As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs7 As New ADODB.Recordset

Dim str1, v_company_code, v_employee_code, v_shift_code As String
Dim v_shift_number, v_flag_batch As Integer
Dim str_date As String

Dim vTglAwal As String
Dim vTglAkhir As String

    dw_work = 0
    i = 0
    j = 0
    
    vTglAwal = Format(Now, "yyyy-MM-dd")
    vTglAkhir = Format(Now, "yyyy-MM-dd")
    
    If CZKEM1.ReadGeneralLogData(vMachineNumber) Then
        While CZKEM1.GetGeneralLogDataStr(vMachineNumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, timeStr)
            j = j + 1
        Wend
    End If
    Call disconnect
    
    Call connect
    If CZKEM1.ReadGeneralLogData(vMachineNumber) Then
        While CZKEM1.GetGeneralLogDataStr(vMachineNumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, timeStr)

            timeStr = Trim(timeStr)

            DoEvents

            str_date = Format(timeStr, "yyyy-MM-dd HH:mm:ss") 'Left(timeStr, Len(timeStr) - 2) & "00"
            
            If rs7.State Then rs7.Close
            SQL = "SELECT * FROM h_log_attendance " & _
                  "WHERE att_date = '" & Format(str_date, "yyyy-MM-dd HH:mm:ss") & "' " & _
                    "AND enrollnumber = '" & Val(dwEnrollNumber) & "' " & _
                    "AND ip_address = '" & FG_IP_ADDRESS & "'"
            rs7.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs7.RecordCount = 0 Then
'                If Format(str_date, "yyyy-MM-dd") < vTglAwal Then vTglAwal = Format(str_date, "yyyy-MM-dd")
'                If Format(str_date, "yyyy-MM-dd") > vTglAkhir Then vTglAkhir = Format(str_date, "yyyy-MM-dd")
            
                SQL = "INSERT INTO h_log_attendance(att_date, ip_address, enrollnumber," & _
                        "verifymode, flag_io, flag_attendance, entry_date) " & _
                      "VALUES( " & _
                        "'" & Format(str_date, "yyyy-MM-dd HH:mm:ss") & "', '" & FG_IP_ADDRESS & "', '" & Val(dwEnrollNumber) & "'," & _
                        "'" & dwVerifyMode & "','" & dwInOutMode & "',0,Now())"
                CnG.Execute SQL
            End If
            rs7.Close
            
            If rs7.State Then rs7.Close
            SQL = "SELECT * FROM m_enroll WHERE enrollnumber = '" & Val(dwEnrollNumber) & "'"
            rs7.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs7.RecordCount = 0 Then
                Dim recFP As New ADODB.Recordset
                Dim name As String, passWord As String, privilege As Long, Enabled As Boolean
                Dim iBackupNumber As Long
                Dim sTmpData As String
                Dim Flag As Long, tmp_len As Long
                            
                If CZKEM1.ReadAllUserID(vMachineNumber) Then
                    If CZKEM1.GetUserInfo(vMachineNumber, dwEnrollNumber, name, passWord, privilege, Enabled) Then
                        For iBackupNumber = 0 To 9
                            If CZKEM1.GetUserTmpStr(CLng(vMachineNumber), CStr(dwEnrollNumber), CLng(iBackupNumber), sTmpData, CLng(tmp_len)) Then
                                'Add to DB
                                SQL = "SELECT * FROM m_enroll WHERE ip_address = '" & FG_IP_ADDRESS & "' " & _
                                        "AND EnrollNumber = '" & dwEnrollNumber & "' " & _
                                        "AND FingerNumber = '" & iBackupNumber & "'"
                                recFP.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                                
                                If recFP.RecordCount = 0 Then
                                    SQL = "INSERT INTO m_enroll (ip_address, EnrollNumber, EnrollName, EMachineNumber," & _
                                            "FingerNumber, Privilige, Password, FPData) " & _
                                          "VALUES (" & _
                                            "'" & FG_IP_ADDRESS & "', '" & dwEnrollNumber & "', '" & Replace(name, "'", "''") & "'," & _
                                            "'" & vMachineNumber & "', '" & iBackupNumber & "', '" & privilege & "'," & _
                                            "'" & passWord & "', '" & sTmpData & "')"
                                    CnG.Execute SQL
                                End If
                                recFP.Close
                            End If
                        Next iBackupNumber
                        
                        DoEvents
                    End If
                End If
            End If
            rs7.Close
            
            i = i + 1
            lbl_progress.Caption = "Downloading...(" & i & " / " & j & ")"
        Wend
        
        Screen.MousePointer = vbDefault
    Else
        download_data_log_typeA = False
    End If

    MousePointer = vbDefault
    If public_int_caller = 0 Then
        MsgBox i & " are successfully downloaded...", vbInformation, headerMSG
        lbl_progress.Caption = "Finish Downloading..."
        
        If vAutoDeleteLog <> 0 Then
            If clear_all_log = False Then
                If Not public_int_caller = 1 Then
                    lbl_progress.Caption = "Error deleting log..."
                Else
                    lbl_progress.Caption = "Delete log successfully..."
                End If
            End If
        End If
    End If

    download_data_log_typeA = True
    If i = 0 Then MsgBox "There is no data...", vbCritical, headerMSG
    Exit Function

capErr:
MsgBox "Error downloading data..." & vbCr & _
    Err.Description, vbCritical, headerMSG
MousePointer = vbDefault
download_data_log_typeA = False
End Function

Public Function download_data_log_typeB() As Boolean
On Error Resume Next

Dim dwEnrollNumber As String 'As Long
Dim dwVerifyMode As Long
Dim dwInOutMode As Long
Dim timeStr As String
Dim i As Long
Dim lng_year As Long
Dim lng_month As Long
Dim lng_day As Long
Dim lng_hour As Long
Dim lng_minute As Long
Dim lng_second As Long
Dim dw_work As Long
Dim j As Long

Dim rs1 As New ADODB.Recordset
Dim rs7 As New ADODB.Recordset

Dim str1, v_company_code, v_employee_code, v_shift_code As String
Dim v_shift_number, v_flag_batch As Integer
Dim str_date As String

Dim vTglAwal As String
Dim vTglAkhir As String

Dim rsabsen As New ADODB.Recordset

    dw_work = 0
    i = 0
    j = 0
    vTglAwal = Format(Now, "yyyy-MM-dd")
    vTglAkhir = Format(Now, "yyyy-MM-dd")
    
    If CZKEM1.ReadAllGLogData(vMachineNumber) Then
        While CZKEM1.SSR_GetGeneralLogData(vMachineNumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, _
        lng_year, lng_month, lng_day, lng_hour, lng_minute, lng_second, dw_work)
            j = j + 1
        Wend
    End If
    Call disconnect
    
    Call connect
    If CZKEM1.ReadAllGLogData(vMachineNumber) Then
        While CZKEM1.SSR_GetGeneralLogData(vMachineNumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, _
        lng_year, lng_month, lng_day, lng_hour, lng_minute, lng_second, dw_work)
                    
            timeStr = Trim(timeStr)
            
            DoEvents
                  
            str_date = _
            Trim(str(lng_year)) & "-" & Right("00" & Trim(str(lng_month)), 2) _
                                & "-" & Right("00" & Trim(str(lng_day)), 2) _
            & " " & Right("00" & Trim(str(lng_hour)), 2) & ":" & Right("00" & Trim(str(lng_minute)), 2) _
                                                        & ":" & Right("00" & Trim(str(lng_second)), 2)
            
            If rs7.State Then rs7.Close
            SQL = "SELECT * FROM h_log_attendance " & _
                  "WHERE att_date = '" & Format(str_date, "yyyy-MM-dd HH:mm:ss") & "' " & _
                    "AND enrollnumber = '" & Val(dwEnrollNumber) & "' " & _
                    "AND ip_address = '" & FG_IP_ADDRESS & "'"
            rs7.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs7.RecordCount = 0 Then
'                If Format(str_date, "yyyy-MM-dd") < vTglAwal Then vTglAwal = Format(str_date, "yyyy-MM-dd")
'                If Format(str_date, "yyyy-MM-dd") > vTglAkhir Then vTglAkhir = Format(str_date, "yyyy-MM-dd")
                
                SQL = "INSERT INTO h_log_attendance(att_date, ip_address, enrollnumber," & _
                        "verifymode, flag_io, flag_attendance, entry_date) " & _
                      "VALUES( " & _
                        "'" & Format(str_date, "yyyy-MM-dd HH:mm:ss") & "', '" & FG_IP_ADDRESS & "', '" & Val(dwEnrollNumber) & "'," & _
                        "'" & dwVerifyMode & "','" & dwInOutMode & "',0,Now())"
                CnG.Execute SQL
            End If
            rs7.Close
            
            If rs7.State Then rs7.Close
            SQL = "SELECT * FROM m_enroll WHERE enrollnumber = '" & Val(dwEnrollNumber) & "'"
            rs7.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs7.RecordCount = 0 Then
                Dim recFP As New ADODB.Recordset
                Dim name As String, passWord As String, privilege As Long, Enabled As Boolean
                Dim iBackupNumber
                Dim sTmpData As String
                Dim Flag As Long, tmp_len As Long
                            
                If CZKEM1.ReadAllUserID(vMachineNumber) Then
                    If CZKEM1.SSR_GetUserInfo(vMachineNumber, dwEnrollNumber, name, passWord, privilege, Enabled) Then
                        For iBackupNumber = 0 To 9
                            If CZKEM1.SSR_GetUserTmpStr(CLng(vMachineNumber), CStr(dwEnrollNumber), CLng(iBackupNumber), sTmpData, CLng(tmp_len)) Then
                                'Add to DB
                                SQL = "SELECT * FROM m_enroll WHERE ip_address = '" & FG_IP_ADDRESS & "' " & _
                                        "AND EnrollNumber = '" & dwEnrollNumber & "' " & _
                                        "AND FingerNumber = '" & iBackupNumber & "'"
                                recFP.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                                
                                If recFP.RecordCount = 0 Then
                                    SQL = "INSERT INTO m_enroll (ip_address, EnrollNumber, EnrollName, EMachineNumber," & _
                                            "FingerNumber, Privilige, Password, FPData) " & _
                                          "VALUES (" & _
                                            "'" & FG_IP_ADDRESS & "', '" & dwEnrollNumber & "', '" & Replace(name, "'", "''") & "'," & _
                                            "'" & vMachineNumber & "', '" & iBackupNumber & "', '" & privilege & "'," & _
                                            "'" & passWord & "', '" & sTmpData & "')"
                                    CnG.Execute SQL
                                End If
                                recFP.Close
                            End If
                        Next iBackupNumber
                        
                        DoEvents
                    End If
                End If
            End If
            rs7.Close
            i = i + 1
            lbl_progress.Caption = "Downloading...(" & i & " / " & j & ")"
        Wend

        Screen.MousePointer = vbDefault
    Else
        download_data_log_typeB = False
    End If
    
    MousePointer = vbDefault
    If public_int_caller = 0 Then
        MsgBox i & " are successfully downloaded...", vbInformation, headerMSG
        lbl_progress.Caption = "Finish Downloading..."
        
        If vAutoDeleteLog <> 0 Then
            If clear_all_log = False Then
                If Not public_int_caller = 1 Then
                    lbl_progress.Caption = "Error deleting log..."
                Else
                    lbl_progress.Caption = "Delete log successfully..."
                End If
            End If
        End If
    End If
    
    download_data_log_typeB = True
    If i = 0 Then MsgBox "There is no data...", vbCritical, headerMSG
    Exit Function

capErr:
MsgBox "Error downloading data..." & vbCr & _
    Err.Description, vbCritical, headerMSG
MousePointer = vbDefault
download_data_log_typeB = False
End Function

Private Function check_auto_log() As Boolean
Dim rs As New ADODB.Recordset
    check_auto_log = False
    
    If rs.State Then rs.Close
    rs.Open "select * from s_device where s_number = 1", CnG, adOpenStatic, adLockReadOnly
    If rs.RecordCount = 1 Then
        If rs.Fields("s_device").Value = 1 Then
            check_auto_log = True
        End If
    Else
        check_auto_log = False
        Exit Function
    End If
    
    If rs.State = 1 Then rs.Close
    rs.Open "select count(*) as rec_count from s_auto_log where left(cast(s_time as time),5)='" _
        & Format(Now, "hh:mm") & "' and ifnull(s_enable,0)=1", CnG, adOpenStatic, adLockReadOnly
    'check_auto_log = IIf(rs.Fields("rec_count").Value >= 1 And mnu_trans_log_att.Enabled = True, True, False)
End Function

Private Sub execCommand(ByVal cmd As String)
    Dim result  As Long
    Dim lPid    As Long
    Dim lHnd    As Long
    Dim lRet    As Long

'    cmd = "cmd /c " & cmd
    
    cmd = Chr(34) & App.Path & "\backup.bat" & Chr(34)
    result = Shell(cmd, vbHide)
    
    lPid = result
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
    End If
End Sub

Private Sub CreateBatchFile(strPath As String, BatchCommands As String)
Dim strOutput As String: strOutput = BatchCommands
    strOutput = Replace(strOutput, "/", vbCrLf)
    Open strPath For Output As #1
    Print #1, strOutput
    Close #1
End Sub


Private Sub createGridKar()
   With LynxGrid2
      .AddColumn "Company", 2000, , , , , , , , , True
      .AddColumn "Employee Code", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 2000, , , , , , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear
                
        If rs.State Then rs.Close
        SQL = "select b.company_name,employee_code,employee_name " & _
                "from m_employee a join m_company b on a.company_code = b.company_code " & _
                "WHERE flag_active <> 0 " & _
                   "AND (employee_code LIKE '%" & txt_employee_code.Text & "%' " & _
                       "OR employee_name LIKE '%" & txt_employee_code.Text & "%')"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid2.AddItem rs!company_name & vbTab & rs!employee_code & vbTab & rs!EMPLOYEE_NAME
                rs.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                txt_employee_code.Text = rs!employee_code
                txt_employee_name.Text = rs!EMPLOYEE_NAME
'                TDBCombo1.SetFocus
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else
            
        End If
        rs.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txt_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txt_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub LynxGrid2_DblClick()
    isiGridKar (2)
End Sub

Private Sub LynxGrid2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        LynxGrid2.Visible = False
    End If
    If KeyAscii = 13 Then
        isiGridKar (2)
    End If
End Sub

Private Sub LynxGrid2_LostFocus()
    LynxGrid2.Visible = False
End Sub

Private Sub txt_employee_code_Change()
    If txt_employee_code.Text = "" Then
        txt_employee_name.Text = ""
    End If
End Sub

Private Sub txt_employee_code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmdBrowse_Click()
    isiGridKar (1)
End Sub

