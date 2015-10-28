VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_etc_backup_restore 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup & Restore"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   Icon            =   "frmBackupRestore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prj_panji.vbButton cmdOK 
      Height          =   450
      Left            =   4380
      TabIndex        =   2
      Top             =   2235
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
      MICON           =   "frmBackupRestore.frx":058A
      PICN            =   "frmBackupRestore.frx":05A6
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
      Caption         =   "Backup And Restore"
      Height          =   1995
      Left            =   2340
      TabIndex        =   0
      Top             =   90
      Width           =   4695
      Begin VB.OptionButton optType 
         Caption         =   "Restore"
         Height          =   195
         Index           =   1
         Left            =   2010
         TabIndex        =   8
         Top             =   390
         Width           =   915
      End
      Begin VB.OptionButton optType 
         Caption         =   "Backup"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   7
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   990
         TabIndex        =   4
         Top             =   810
         Width           =   3015
      End
      Begin prj_panji.vbButton btnBrowse 
         Height          =   315
         Left            =   4080
         TabIndex        =   6
         Top             =   810
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   "..."
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
         MICON           =   "frmBackupRestore.frx":1638
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama File :"
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   870
         Width           =   885
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Pilih Nama File dan Lokasi File Backup/Restore"
         Height          =   405
         Left            =   150
         TabIndex        =   1
         Top             =   1410
         Width           =   3945
      End
   End
   Begin prj_panji.vbButton cmdCancel 
      Height          =   450
      Left            =   5730
      TabIndex        =   3
      Top             =   2235
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
      BTYPE           =   14
      TX              =   "&Batal"
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
      MICON           =   "frmBackupRestore.frx":1654
      PICN            =   "frmBackupRestore.frx":1670
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2340
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save File To"
      Filter          =   "*.sql"
      InitDir         =   "C:\"
   End
   Begin VB.Image Image2 
      Height          =   2850
      Left            =   0
      Picture         =   "frmBackupRestore.frx":2702
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2265
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   2280
      X2              =   2280
      Y1              =   -30
      Y2              =   3390
   End
   Begin VB.Image Image1 
      Height          =   3780
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2265
   End
End
Attribute VB_Name = "frm_etc_backup_restore"
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

Private Sub btnBrowse_Click()
    If optType(0) Then
        CommonDialog1.ShowSave
    Else
        CommonDialog1.ShowOpen
    End If
    
    If Len(CommonDialog1.FileName) > 0 And Right(CommonDialog1.FileName, 4) <> ".sql" Then
        txtFileName.Text = CommonDialog1.FileName & ".sql"
    ElseIf Len(CommonDialog1.FileName) > 0 And Right(CommonDialog1.FileName, 4) = ".sql" Then
        txtFileName.Text = CommonDialog1.FileName
    End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err
Dim cmd As String
Dim location As String
Dim result As Double

    If Len(txtFileName.Text) = 0 Then
        MsgBox "Nama File Tidak Sesuai...", vbExclamation, headerMSG
        Exit Sub
    End If

    location = Chr$(34) & Trim(txtFileName.Text) & Chr$(34)
    
    Screen.MousePointer = vbHourglass
    DoEvents
    If optType(0) Then
        cmd = Chr(34) & App.Path & "\mysql\bin\mysqldump" & Chr(34) & " -h" & ServerDB & " -u" & UserDB & " -p" & passDB & " --routines --comments " & nmDB & " > " & location
        Call CreateBatchFile(App.Path & "\backup.bat", cmd)
    Else
        cmd = Chr(34) & App.Path & "\mysql\bin\mysql" & Chr(34) & " -h" & ServerDB & " -u" & UserDB & " -p" & passDB & " --comments " & nmDB & " < " & location
        Call CreateBatchFile(App.Path & "\restore.bat", cmd)
    End If
    
    Call execCommand(cmd)
    
    Screen.MousePointer = vbDefault
    
    If optType(0) Then
        MsgBox "Proses Backup Selesai...", vbInformation, headerMSG
    Else
        MsgBox "Proses Restore Selesai...", vbInformation, headerMSG
    End If
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub Form_Load()
    optType(0) = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_etc_backup_restore = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CreateBatchFile(strPath As String, BatchCommands As String)
Dim strOutput As String: strOutput = BatchCommands
    strOutput = Replace(strOutput, "/", vbCrLf)
    Open strPath For Output As #1
    Print #1, strOutput
    Close #1
End Sub

Private Sub execCommand(ByVal cmd As String)
    Dim result  As Long
    Dim lPid    As Long
    Dim lHnd    As Long
    Dim lRet    As Long

'    cmd = "cmd /c " & cmd
    
    If optType(0) Then
        cmd = Chr(34) & App.Path & "\backup.bat" & Chr(34)
        result = Shell(cmd, vbHide)
    Else
        cmd = Chr(34) & App.Path & "\restore.bat" & Chr(34)
        result = Shell(cmd, vbHide)
    End If
    
    lPid = result
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
    End If
End Sub
