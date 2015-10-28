VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_etc_backup_restore_device 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup & Restore Finger Data"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   Icon            =   "frm_backup_restore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   345
      Left            =   3150
      OleObjectBlob   =   "frm_backup_restore.frx":058A
      TabIndex        =   8
      Top             =   2340
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Frame Frame1 
      Caption         =   "Backup And Restore"
      Height          =   1995
      Left            =   2340
      TabIndex        =   0
      Top             =   90
      Width           =   4695
      Begin VB.TextBox txt_device_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   870
         Width           =   2535
      End
      Begin VB.OptionButton optType 
         Caption         =   "Restore"
         Height          =   195
         Index           =   1
         Left            =   2010
         TabIndex        =   4
         Top             =   390
         Width           =   915
      End
      Begin VB.OptionButton optType 
         Caption         =   "Backup"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   390
         Value           =   -1  'True
         Width           =   915
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_device 
         Height          =   375
         Left            =   690
         OleObjectBlob   =   "frm_backup_restore.frx":05AE
         TabIndex        =   6
         Top             =   870
         Width           =   1335
      End
      Begin VB.TextBox txt_port 
         Height          =   285
         Left            =   3150
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   690
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DARI :"
         Height          =   255
         Left            =   -240
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   1620
         Width           =   4425
      End
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
   Begin prj_tpc.vbButton cmdOK 
      Height          =   450
      Left            =   4380
      TabIndex        =   9
      Top             =   2250
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
      MICON           =   "frm_backup_restore.frx":299F
      PICN            =   "frm_backup_restore.frx":29BB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdCancel 
      Height          =   450
      Left            =   5730
      TabIndex        =   10
      Top             =   2250
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
      MICON           =   "frm_backup_restore.frx":3A4D
      PICN            =   "frm_backup_restore.frx":3A69
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
      Height          =   2850
      Left            =   0
      Picture         =   "frm_backup_restore.frx":4AFB
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
Attribute VB_Name = "frm_etc_backup_restore_device"
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

Private Function connect() As Boolean
    If cn_fp Then
        CZKEM1.EnableDevice vMachineNumber, True
        CZKEM1.disconnect
    End If
    
    cn_fp = CZKEM1.Connect_Net(TDBCombo_device.Text, CLng(txt_port.Text))
    
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
                                
                                lblCaption.Caption = "Backup Data Ke - " & vFlagSuccess
                            End If
                        Next iBackupNumber
                        DoEvents
                        
                    Wend
                                    
                    MsgBox vFlagSuccess & " Data Berhasil Di Backup...", vbInformation, headerMSG
                ElseIf alg_ver = "" Then
                    While CZKEM1.GetAllUserInfo(vMachineNumber, dwEnrollNmber, name, passWord, privilege, Enabled)
                        For iBackupNumber = 0 To 9
                            If CZKEM1.GetUserTmpStr(CLng(vMachineNumber), dwEnrollNmber, CLng(iBackupNumber), sTmpData, CLng(tmp_len)) Then
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
                                
                                lblCaption.Caption = "Backup Data Ke - " & vFlagSuccess
                            End If
                        Next iBackupNumber
                        DoEvents
                        
                    Wend
                                    
                    MsgBox vFlagSuccess & " Data Berhasil Di Backup...", vbInformation, headerMSG
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
                                lblCaption.Caption = "Backup Data Ke - " & vFlagSuccess
                            End If
                        Next iBackupNumber
                        DoEvents
                    Wend
                    
                    MsgBox vFlagSuccess & " Data Berhasil Di Backup...", vbInformation, headerMSG
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
                    If Not IsNumeric(alg_ver) Then alg_ver = ""
            
                    If CZKEM1.ReadAllUserID(vMachineNumber) Then
                        If alg_ver = "9" And (CZKEM1.IsTFTMachine(vMachineNumber) = False) Then 'Alg 9
                            If CZKEM1.SetUserTmpStr(vMachineNumber, iEnrollNumber, iBackupNumber, sEnrollData) Then
                                vFlagSuccess = vFlagSuccess + 1
                            Else
                                vFlagFailed = vFlagFailed + 1
                            End If
                        ElseIf alg_ver = "" Then
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
                    
                    lblCaption.Caption = "Restore Data Ke - " & vFlagSuccess
                recFP.MoveNext
                DoEvents
                Wend
                
                MsgBox vFlagSuccess & " Data Berhasil Di Restore..." & Chr(13) & _
                        vFlagFailed & " Data Gagal Di Restore...", vbInformation, headerMSG
                        
                lblCaption.Caption = ""
            End If
            recFP.Close
        End If
        Call disconnect
    Else
        MsgBox "Tidak Dapat Menghubungkan ke Mesin Fingerprint...", vbExclamation, headerMSG
        Exit Sub
    End If
    Me.MousePointer = vbNormal
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub Form_Load()
    Call load_data_device
    
    vMachineNumber = 1
        
    optType(0) = True
End Sub

Private Sub load_data_device()
    If rsDevice.State = 1 Then rsDevice.Close
    SQL = "select * from m_device order by ip_address"
    rsDevice.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_device.RowSource = rsDevice
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_etc_backup_restore_device = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub optType_Click(Index As Integer)
    If optType(0) Then
        lblName.Caption = "DARI :"
    Else
        lblName.Caption = "KE :"
    End If
End Sub

Private Sub TDBCombo_device_ItemChange()
    If TDBCombo_device.ApproxCount > 0 Then
        TDBCombo_device.Text = TDBCombo_device.Columns("ip_address").Value
        txt_device_name.Text = TDBCombo_device.Columns("name").Value
        txt_port.Text = TDBCombo_device.Columns("port").Value
    End If
End Sub

'Private Sub getUserTemplate()
'Dim iEnrollNumber As Long, iBackupNumber As Long
'Dim sTmpData As String
'Dim Flag As Long, tmp_len As Long
'
'    tmp_len = 1024 * 20 '//Fixed Value
'    txtEvent.Text = ""
'    iEnrollNumber = Pin.Text
'    iBackupNumber = PinIndex.Text
'    x = CZKEM1.EnableDevice(iMachineNumber, False)
'
'    'Only TFT screen devices with firmware version Ver 6.60 version later support function "GetUserTmpExStr" and "GetUserTmpEx".
'    'While you are using 9.0 fingerprint arithmetic and your device's firmware version is under ver6.60,you should use the functions "SSR_GetUserTmp" or
'    '"SSR_GetUserTmpStr" instead of "GetUserTmpExStr" or "GetUserTmpEx" in order to download the fingerprint templates.
'    If CZKEM1.GetUserTmpExStr(vMachineNumber, iEnrollNumber, iBackupNumber, Flag, sTmpData, CLng(tmp_len)) Then
'
'        'Add to DB
'        recFP.Open "select * from fptable", connFP, adOpenKeyset, adLockOptimistic
'        recFP.AddNew
'        recFP.Fields("EMachineNumber") = vMachineNumber
'        recFP.Fields("EnrollNumber") = iEnrollNumber
'        recFP.Fields("FingerNumber") = iBackupNumber
'        recFP.Fields("Template") = sTmpData
'        recFP.Update
'        recFP.Close
'        txtEvent.Text = "Saved On Database"
'        lblInfo.Caption = "SSR_GetUserTmpStr OK"
'    Else
'        MsgBox "Gagal Download"
'    End If
'    x = CZKEM1.EnableDevice(iMachineNumber, True)
'End Sub
