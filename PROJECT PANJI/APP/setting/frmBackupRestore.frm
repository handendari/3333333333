VERSION 5.00
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_etc_backup_restore 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup & Restore"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   Icon            =   "frmBackupRestore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Backup And Restore"
      Height          =   3285
      Left            =   2340
      TabIndex        =   0
      Top             =   90
      Width           =   4695
      Begin prj_panji.vbButton btnBrowse 
         Height          =   315
         Left            =   4050
         TabIndex        =   25
         Top             =   2070
         Width           =   375
         _ExtentX        =   661
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
         MPTR            =   0
         MICON           =   "frmBackupRestore.frx":058A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame fraAutomatic 
         Height          =   1275
         Left            =   960
         TabIndex        =   10
         Top             =   720
         Width           =   3105
         Begin VB.ComboBox cboAuto 
            Height          =   315
            ItemData        =   "frmBackupRestore.frx":05A6
            Left            =   90
            List            =   "frmBackupRestore.frx":05B3
            TabIndex        =   11
            Top             =   180
            Width           =   2925
         End
         Begin VB.Frame fraAuto 
            Height          =   615
            Index           =   2
            Left            =   90
            TabIndex        =   19
            Top             =   480
            Width           =   2925
            Begin MSComCtl2.DTPicker DTPicker_monthly 
               Height          =   315
               Left            =   540
               TabIndex        =   22
               Top             =   210
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd-MMM"
               Format          =   95289347
               CurrentDate     =   41351
            End
            Begin TDBTime6Ctl.TDBTime monthly_time 
               Height          =   285
               Left            =   2220
               TabIndex        =   20
               Top             =   210
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   503
               Caption         =   "frmBackupRestore.frx":05CF
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmBackupRestore.frx":063B
               Spin            =   "frmBackupRestore.frx":068B
               AlignHorizontal =   2
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "hh:nn"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "hh:nn"
               HighlightText   =   0
               Hour12Mode      =   1
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxTime         =   0.99999
               MidnightMode    =   0
               MinTime         =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "00:00"
               ValidateMode    =   0
               ValueVT         =   7
               Value           =   0
            End
            Begin VB.Label Label4 
               Caption         =   "DATE"
               Height          =   195
               Left            =   60
               TabIndex        =   21
               Top             =   270
               Width           =   525
            End
         End
         Begin VB.Frame fraAuto 
            Height          =   615
            Index           =   1
            Left            =   90
            TabIndex        =   15
            Top             =   480
            Width           =   2925
            Begin VB.ComboBox cboWeek 
               Height          =   315
               ItemData        =   "frmBackupRestore.frx":06B3
               Left            =   570
               List            =   "frmBackupRestore.frx":06CC
               TabIndex        =   18
               Top             =   210
               Width           =   1635
            End
            Begin TDBTime6Ctl.TDBTime weekly_time 
               Height          =   285
               Left            =   2220
               TabIndex        =   16
               Top             =   210
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   503
               Caption         =   "frmBackupRestore.frx":0710
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmBackupRestore.frx":077C
               Spin            =   "frmBackupRestore.frx":07CC
               AlignHorizontal =   2
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "hh:nn"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "hh:nn"
               HighlightText   =   0
               Hour12Mode      =   1
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxTime         =   0.99999
               MidnightMode    =   0
               MinTime         =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "00:00"
               ValidateMode    =   0
               ValueVT         =   7
               Value           =   0
            End
            Begin VB.Label Label2 
               Caption         =   "DAY :"
               Height          =   195
               Left            =   60
               TabIndex        =   17
               Top             =   240
               Width           =   525
            End
         End
         Begin VB.Frame fraAuto 
            Height          =   615
            Index           =   0
            Left            =   90
            TabIndex        =   12
            Top             =   480
            Width           =   2925
            Begin TDBTime6Ctl.TDBTime daily_time 
               Height          =   285
               Left            =   1290
               TabIndex        =   13
               Top             =   210
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   503
               Caption         =   "frmBackupRestore.frx":07F4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmBackupRestore.frx":0860
               Spin            =   "frmBackupRestore.frx":08B0
               AlignHorizontal =   2
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "hh:nn"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "hh:nn"
               HighlightText   =   0
               Hour12Mode      =   1
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxTime         =   0.99999
               MidnightMode    =   0
               MinTime         =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "00:00"
               ValidateMode    =   0
               ValueVT         =   7
               Value           =   0
            End
            Begin VB.Label Label8 
               Caption         =   "TIME :"
               Height          =   195
               Left            =   540
               TabIndex        =   14
               Top             =   240
               Width           =   765
            End
         End
      End
      Begin VB.Frame fra_type 
         Height          =   555
         Left            =   960
         TabIndex        =   7
         Top             =   180
         Width           =   3105
         Begin VB.OptionButton optType 
            Caption         =   "Manual"
            Height          =   255
            Index           =   1
            Left            =   1770
            TabIndex        =   9
            Top             =   210
            Width           =   975
         End
         Begin VB.OptionButton optType 
            Caption         =   "Automatic"
            Height          =   255
            Index           =   0
            Left            =   330
            TabIndex        =   8
            Top             =   210
            Width           =   1095
         End
      End
      Begin VB.Frame fraManual 
         Height          =   555
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   3105
         Begin VB.OptionButton optManual 
            Caption         =   "Backup"
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   6
            Top             =   210
            Width           =   945
         End
         Begin VB.OptionButton optManual 
            Caption         =   "Restore"
            Height          =   225
            Index           =   1
            Left            =   1770
            TabIndex        =   5
            Top             =   210
            Width           =   915
         End
      End
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   990
         TabIndex        =   2
         Top             =   2070
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "File Name :"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   2130
         Width           =   885
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Choose Filename And Location File Backup/Restore"
         Height          =   405
         Left            =   150
         TabIndex        =   1
         Top             =   2640
         Width           =   3945
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2340
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save File To"
      Filter          =   "*.sql"
      InitDir         =   "C:\"
   End
   Begin prj_panji.vbButton cmdOK 
      Height          =   450
      Left            =   4380
      TabIndex        =   23
      Top             =   3480
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
      MICON           =   "frmBackupRestore.frx":08D8
      PICN            =   "frmBackupRestore.frx":08F4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.vbButton cmdCancel 
      Height          =   450
      Left            =   5730
      TabIndex        =   24
      Top             =   3480
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
      MICON           =   "frmBackupRestore.frx":1986
      PICN            =   "frmBackupRestore.frx":19A2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image2 
      Height          =   4020
      Left            =   0
      Picture         =   "frmBackupRestore.frx":2A34
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
      Height          =   4050
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

Dim rsAutoBak As New ADODB.Recordset

Private Sub btnBrowse_Click()
    If optManual(0) Then
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

Private Sub cboAuto_Click()
    If cboAuto.ListIndex = 0 Then
        fraAuto(0).Visible = True
        fraAuto(1).Visible = False
        fraAuto(2).Visible = False
    ElseIf cboAuto.ListIndex = 1 Then
        fraAuto(0).Visible = False
        fraAuto(1).Visible = True
        fraAuto(2).Visible = False
    ElseIf cboAuto.ListIndex = 2 Then
        fraAuto(0).Visible = False
        fraAuto(1).Visible = False
        fraAuto(2).Visible = True
    End If
End Sub

Private Sub cmdOK_Click()
'On Error GoTo Err
Dim cmd As String
Dim location As String
Dim result As Double
    
    If Len(txtFileName.Text) = 0 Then
        MsgBox "Invalid File Name....!!!", vbExclamation, headerMSG
        Exit Sub
    End If
    
    If optType(1).Value Then
        location = Chr$(34) & Trim(txtFileName.Text) & Chr$(34)
        
        Screen.MousePointer = vbHourglass
        DoEvents
        If optManual(0) Then
            cmd = Chr(34) & App.Path & "\mysql\bin\mysqldump" & Chr(34) & " -h" & ServerDB & " -u" & UserDB & " -p" & passDB & " --routines --comments " & nmDB & " > " & location
            Call CreateBatchFile(App.Path & "\backup.bat", cmd)
        Else
            cmd = Chr(34) & App.Path & "\mysql\bin\mysql" & Chr(34) & " -h" & ServerDB & " -u" & UserDB & " -p" & passDB & " --comments " & nmDB & " < " & location
            Call CreateBatchFile(App.Path & "\restore.bat", cmd)
        End If
        
        Call execCommand(cmd)
        
        Screen.MousePointer = vbDefault
        
        If optManual(0) Then
            MsgBox "Backup Process Completed!", vbInformation, headerMSG
        Else
            MsgBox "Restore Process Completed!", vbInformation, headerMSG
        End If
    Else
        Dim vDailyTime As String
        Dim vWeeklyTime As String
        Dim vMonthlyTime As String
        
        vDailyTime = Format(Now, "yyyy-MM-dd") & " " & Format(daily_time, "hh:mm:00")
        vWeeklyTime = Format(Now, "yyyy-MM-dd") & " " & Format(weekly_time, "hh:mm:00")
        vMonthlyTime = Format(DTPicker_monthly.Value, "yyyy-MM-dd") & " " & Format(monthly_time, "hh:mm:00")
        
        SQL = "DELETE FROM s_auto_bak"
        CnG.Execute SQL
        
        SQL = "INSERT INTO s_auto_bak (location,flag_auto,flag_type,time_daily,day_code," & _
                "day_name,time_weekly,time_monthly)" & _
              "VALUES ( " & _
                "'" & Replace(Trim(txtFileName.Text), "\", "\\") & "','" & cboAuto.ListIndex & "','" & IIf(optType(0).Value, 0, 1) & "'," & _
                "'" & vDailyTime & "','" & cboWeek.ListIndex & "','" & cboWeek.Text & "','" & vWeeklyTime & "','" & vMonthlyTime & "')"
        CnG.Execute SQL
        
        MsgBox "Save Succesfully", vbInformation, headerMSG
    End If
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub Form_Load()
    
    optManual(0) = True
    optType(0) = True
    
    cboAuto.ListIndex = 0
    daily_time.Value = Format(Now, "hh:mm")
    cboWeek.ListIndex = 0
    weekly_time.Value = Format(Now, "hh:mm")
    DTPicker_monthly.Value = Format(Now, "dd-MMM")
    monthly_time.Value = Format(Now, "hh:mm")
    
    Call load_data
    Call set_edit_data
    
    Call load_data_user_access(Me)
    cmdOK.Enabled = blnUser_Add
End Sub

Private Sub load_data()
    If rsAutoBak.State Then rsAutoBak.Close
    SQL = "SELECT * FROM s_auto_bak"
    rsAutoBak.Open SQL, CnG, adOpenForwardOnly
    
End Sub

Private Sub set_edit_data()
Dim vFlagType As Integer
    If rsAutoBak.RecordCount > 0 Then
        With rsAutoBak
            
            vFlagType = .Fields("flag_type").Value
            optType(0).Value = IIf(vFlagType = 0, 1, 0)
            optType(1).Value = IIf(vFlagType = 0, 0, 1)
            
            cboAuto.ListIndex = .Fields("flag_auto").Value
            
            daily_time.Value = Format(.Fields("time_daily").Value, "hh:mm")
            cboWeek.ListIndex = .Fields("day_code").Value
            cboWeek.Text = .Fields("day_name").Value
            weekly_time.Value = Format(.Fields("time_weekly").Value, "hh:mm")
            DTPicker_monthly.Value = Format(.Fields("time_monthly").Value, "dd-MMM")
            txtFileName.Text = .Fields("location").Value
        End With
    End If
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
    
    If optManual(0) Then
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

Private Sub optType_Click(Index As Integer)
    If Index = 0 Then
        fraAutomatic.Visible = True
        fraManual.Visible = False
    Else
        fraAutomatic.Visible = False
        fraManual.Visible = True
    End If
End Sub
