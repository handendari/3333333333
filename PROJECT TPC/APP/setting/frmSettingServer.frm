VERSION 5.00
Begin VB.Form frm_etc_db_config 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Configuration"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "frmSettingServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prj_tpc.vbButton cmdConnect 
      Height          =   450
      Left            =   4080
      TabIndex        =   12
      Top             =   3315
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   794
      BTYPE           =   14
      TX              =   "&Set Default"
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
      MICON           =   "frmSettingServer.frx":058A
      PICN            =   "frmSettingServer.frx":05A6
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
      Caption         =   "Database Configuration"
      Height          =   3135
      Left            =   2340
      TabIndex        =   5
      Top             =   90
      Width           =   4425
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmSettingServer.frx":1638
         Left            =   930
         List            =   "frmSettingServer.frx":1642
         TabIndex        =   15
         Text            =   "SERVER"
         Top             =   300
         Width           =   2685
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   930
         TabIndex        =   1
         Top             =   1020
         Width           =   945
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   930
         TabIndex        =   2
         Top             =   1350
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   930
         PasswordChar    =   "="
         TabIndex        =   4
         Top             =   2010
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   930
         TabIndex        =   3
         Top             =   1680
         Width           =   1995
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   930
         TabIndex        =   0
         Top             =   690
         Width           =   2655
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Port  : "
         Height          =   255
         Left            =   270
         TabIndex        =   11
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Database  : "
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   1380
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password  : "
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User ID :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1710
         Width           =   765
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter The Database Configuration Correctly"
         Height          =   405
         Left            =   150
         TabIndex        =   7
         Top             =   2550
         Width           =   3435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Host  : "
         Height          =   255
         Left            =   270
         TabIndex        =   6
         Top             =   720
         Width           =   645
      End
   End
   Begin prj_tpc.vbButton cmdCancel 
      Height          =   450
      Left            =   5460
      TabIndex        =   13
      Top             =   3315
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
      MICON           =   "frmSettingServer.frx":1655
      PICN            =   "frmSettingServer.frx":1671
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdSave 
      Height          =   450
      Left            =   2370
      TabIndex        =   14
      Top             =   3300
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
      BTYPE           =   14
      TX              =   "&Save"
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
      MICON           =   "frmSettingServer.frx":2703
      PICN            =   "frmSettingServer.frx":271F
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
      Height          =   3870
      Left            =   0
      Picture         =   "frmSettingServer.frx":37B1
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
Attribute VB_Name = "frm_etc_db_config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'Dim sqlobj As New SQLDMO.Application
Dim clskonek As New clsConnect
Dim vIndex As Integer

Private Sub check_koneksi()
    If IsNull(ServerDB) Or ServerDB = "" Then
        cboType.ListIndex = 0
        
        Text1.Text = "localhost"
        Text5.Text = "3306"
        Text4.Text = "db_tpc"
        Text2.Text = "root"
        Text3.Text = ""
    Else
        cboType.ListIndex = TypeDB
        
        Text1.Text = ServerDB
        Text5.Text = PortDB
        Text4.Text = nmDB
        Text2.Text = UserDB
        Text3.Text = ""
    End If
End Sub

Private Sub cboType_Click()
    vIndex = IIf(cboType.ListIndex = 0, 0, 1)
        
    If GetInitEntry(vIndex, "CONFIG", "A") <> "" Then
        Text1.Text = DecryptINI(GetInitEntry(vIndex, "CONFIG", "A"), pEncryptionPassword)
        Text2.Text = DecryptINI(GetInitEntry(vIndex, "CONFIG", "B"), pEncryptionPassword)
        Text3.Text = ""
        Text4.Text = DecryptINI(GetInitEntry(vIndex, "CONFIG", "D"), pEncryptionPassword)
        Text5.Text = DecryptINI(GetInitEntry(vIndex, "CONFIG", "E"), pEncryptionPassword)
    End If
End Sub

Private Sub Form_Load()
    Call check_koneksi
        
'    Call load_data_user_access(Me)
'    cmdConnect.Enabled = blnUser_Add
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_etc_db_config = Nothing
End Sub

Private Sub Option1_Click()
    Text2.Enabled = True
    Text3.Enabled = True
End Sub

Private Sub Option2_Click()
    Text2.Enabled = False
    Text3.Enabled = False
End Sub

Private Sub cmdConnect_Click()
    vLoadMode = 2
    
    vIndex = IIf(cboType.ListIndex = 0, 0, 1)
    
    'SetInitEntry "ConnectionString", "String", strConn
    SetInitEntry vIndex, "CONFIG", "A", EncryptINI(Text1.Text, pEncryptionPassword)
    SetInitEntry vIndex, "CONFIG", "B", EncryptINI(Text2.Text, pEncryptionPassword)
    If Text3.Text <> "" Then
        SetInitEntry vIndex, "CONFIG", "C", EncryptINI(Text3.Text, pEncryptionPassword)
    End If
    SetInitEntry vIndex, "CONFIG", "D", EncryptINI(Text4.Text, pEncryptionPassword)
    SetInitEntry vIndex, "CONFIG", "E", EncryptINI(Text5.Text, pEncryptionPassword)
    
    SetInitEntry 3, "CONFIG", "A", EncryptINI(cboType.ListIndex, pEncryptionPassword)
    
    If clskonek.Koneksi = False Then
'        MsgBox "Error Connecting Database!!" & Chr(13) & _
'            "Please Check Database Configuration!", vbCritical, "Att & Payroll System"
        Exit Sub
    Else
        MsgBox "Connection Database Successfully..." & Chr(13) & _
            "Please Exit Aplication First...", vbInformation, "Att & Payroll System"
        End
    End If
    
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim vIndex As Integer
    vLoadMode = 2
    
    vIndex = IIf(cboType.ListIndex = 0, 0, 1)
    
    'SetInitEntry "ConnectionString", "String", strConn
    SetInitEntry vIndex, "CONFIG", "A", EncryptINI(Text1.Text, pEncryptionPassword)
    SetInitEntry vIndex, "CONFIG", "B", EncryptINI(Text2.Text, pEncryptionPassword)
    If Text3.Text <> "" Then
        SetInitEntry vIndex, "CONFIG", "C", EncryptINI(Text3.Text, pEncryptionPassword)
    End If
    SetInitEntry vIndex, "CONFIG", "D", EncryptINI(Text4.Text, pEncryptionPassword)
    SetInitEntry vIndex, "CONFIG", "E", EncryptINI(Text5.Text, pEncryptionPassword)
    
    MsgBox "Save Successfully...", vbInformation, headerMSG
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdConnect_Click
End Sub


