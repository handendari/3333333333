VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_etc_register_expired 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REGISTER THIS PRODUCT"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7095
   Icon            =   "frm_etc_register_expired.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4545
      Left            =   270
      TabIndex        =   3
      Top             =   1950
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   8017
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "REQUEST KEY"
      TabPicture(0)   =   "frm_etc_register_expired.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmd_request_key"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "REGISTER KEY"
      TabPicture(1)   =   "frm_etc_register_expired.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd_register"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   2835
         Left            =   -74940
         TabIndex        =   13
         Top             =   540
         Width           =   6285
         Begin VB.TextBox txt_register_key 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1260
            Width           =   3435
         End
         Begin prj_panji.vbButton cmd_reg_browse 
            Height          =   285
            Left            =   5520
            TabIndex        =   18
            Top             =   1260
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
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
            MICON           =   "frm_etc_register_expired.frx":05C2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "ACTIVATION KEY FILE*"
            Height          =   195
            Left            =   210
            TabIndex        =   15
            Top             =   1290
            Width           =   1755
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2835
         Left            =   60
         TabIndex        =   4
         Top             =   540
         Width           =   6285
         Begin prj_panji.vbButton cmd_req_browse 
            Height          =   285
            Left            =   5580
            TabIndex        =   17
            Top             =   2220
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
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
            MICON           =   "frm_etc_register_expired.frx":05DE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txt_registration_name 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1920
            TabIndex        =   7
            Top             =   450
            Width           =   3615
         End
         Begin VB.TextBox txt_company_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1170
            Width           =   3615
         End
         Begin VB.TextBox txt_generate_key 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   2220
            Width           =   3615
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_company 
            Height          =   375
            Left            =   1920
            OleObjectBlob   =   "frm_etc_register_expired.frx":05FA
            TabIndex        =   8
            Top             =   810
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "NAMA REGISTER*"
            Height          =   195
            Left            =   480
            TabIndex        =   11
            Top             =   480
            Width           =   1395
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "PERUSAHAAN*"
            Height          =   195
            Left            =   690
            TabIndex        =   10
            Top             =   870
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "FILE GENERATE KEY*"
            Height          =   195
            Left            =   180
            TabIndex        =   9
            Top             =   2250
            Width           =   1680
         End
      End
      Begin prj_panji.vbButton cmd_request_key 
         Height          =   705
         Left            =   4560
         TabIndex        =   12
         Top             =   3570
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Req. Key"
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
         MICON           =   "frm_etc_register_expired.frx":2560
         PICN            =   "frm_etc_register_expired.frx":257C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_register 
         Height          =   705
         Left            =   -70380
         TabIndex        =   16
         Top             =   3570
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Register"
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
         MICON           =   "frm_etc_register_expired.frx":360E
         PICN            =   "frm_etc_register_expired.frx":362A
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
   Begin VB.CheckBox Check1 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   5490
      Visible         =   0   'False
      Width           =   375
   End
   Begin prj_panji.vbButton CmdExit 
      Height          =   705
      Left            =   5730
      TabIndex        =   2
      Top             =   6540
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
      MICON           =   "frm_etc_register_expired.frx":46BC
      PICN            =   "frm_etc_register_expired.frx":46D8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   6450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl_license 
      Alignment       =   2  'Center
      Caption         =   "This product has no license"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   360
      Picture         =   "frm_etc_register_expired.frx":576A
      Stretch         =   -1  'True
      Top             =   270
      Width           =   6240
   End
End
Attribute VB_Name = "frm_etc_register_expired"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Dim rsa As New clsRSA
Dim rsComp As New ADODB.Recordset
Dim h As New HDSN
Dim v_no As Integer

Dim v_id As String
Dim v_name As String
Dim v_company As String
Dim v_hardisk As String
Dim v_tgl_req As String
Dim v_tgl_reg As String
Dim v_tgl_exp As String
Dim v_unlimited As String

Dim clsReg As New clsCheckRegister

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmd_register_Click()
'On Error GoTo Err

    If txt_register_key.Text <> "" Then
        Open txt_register_key.Text For Input As #1
        Do While Not EOF(1)
            Input #1, v_id, v_name, v_company, v_hardisk, v_tgl_req, v_tgl_reg, v_tgl_exp, v_unlimited
        Loop
        Close #1
    End If
    
    v_id = rsa.Decrypt(rsa.PrivateKey, v_id)
    
    If rs.State Then rs.Close
    SQL = "SELECT * FROM s_about WHERE id = '" & v_id & "'"
    rs.Open SQL, CnG, adOpenForwardOnly
    
    If rs.RecordCount > 0 Then
        v_hardisk = rsa.Decrypt(rsa.PrivateKey, v_hardisk)
        
        If "SSD-" & rsa.Decrypt(rsa.PrivateKey, rs!serial) <> v_hardisk Then
            MsgBox "File Registrasi Tidak Sesuai...", vbCritical, headerMSG
            rs.Close
            Exit Sub
        Else
            SQL = "UPDATE s_about SET register = '" & v_name & "'," & _
                    "company = '" & v_company & "'," & _
                    "tgl1 = '" & v_tgl_req & "'," & _
                    "tgl2 = '" & v_tgl_reg & "'," & _
                    "tgl3 = '" & v_tgl_exp & "', " & _
                    "ul = '" & v_unlimited & "' " & _
                  "WHERE id = '" & v_id & "'"
            CnG.Execute SQL
            
            MsgBox "Registrasi Berhasil..." & Chr(13) & _
                    "Jalankan Ulang Aplikasi...", vbInformation, headerMSG
            End
        End If
    Else
        MsgBox "File Registrasi Tidak Sesuai...", vbCritical, headerMSG
        rs.Close
        Exit Sub
    End If
    rs.Close
    
    Exit Sub
    
Err:
MsgBox "File Registrasi Tidak Sesuai...", vbExclamation, headerMSG
End Sub

Private Function get_add_str(ByVal str1 As String) As String
Dim str2 As String

str2 = "                  : "
get_add_str = str1 & Right(str2, 15 - Len(str1))
End Function

Private Function get_finger_print() As String
Dim h As New HDSN
Dim str1 As String
Dim rs1 As New ADODB.Recordset

get_finger_print = h.GetSerialNumber
'(h.GetModelNumber & "" & h.GetSerialNumber)
End Function

Private Sub Form_Load()
Dim rs2 As New ADODB.Recordset

    SSTab1.Tab = 0
    
    rsa.m_PrivateKey.Key = pKunci1
    rsa.m_PublicKey.Key = pKunci2
    
    rsa.m_PrivateKey.Length = Len(pKunci1)
    rsa.m_PublicKey.Length = Len(pKunci2)
    
    Call load_data_company
    lbl_license = "This product has expired license"

End Sub

Private Function get_license_key() As String
Dim str_l As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

rs.Open "select * from m_license order by drive_id asc", CnG, adOpenStatic, adLockReadOnly

While Not rs.EOF

    h.CurrentDrive = Val(rs.Fields("drive_id").Value)
'    str_l = (h.GetModelNumber & "" & h.GetSerialNumber)
    str_l = h.GetSerialNumber
    If rs.Fields("license_number").Value = RC4DeCryptASC(str_l, pEncryptionPassword) Then
        get_license_key = rs.Fields("license_number").Value
        Exit Function
    Else
        get_license_key = ""
    End If
    
    rs.MoveNext
Wend
End Function

Private Function OpenFile(Mode As Integer, str_filename As String) As Boolean
On Error GoTo Err

If Mode = 0 Then Open str_filename For Input As #1
If Mode = 1 Then Open str_filename For Output As #1
OpenFile = True

Exit Function
Err:
    OpenFile = False
    MsgBox Err.Description
End Function

Private Function read_file(ByVal spath As String) As String
Dim mFile1 As New CFileSys

On Error GoTo Err

Dim str_data, str1, str2 As String, i As Integer
Dim intLine As Integer

str1 = ""

'If Not OpenFile(0, spath) Then GoTo err
Open spath For Input As #1
While Not EOF(1)
    Line Input #1, str_data
    If Not Trim("" & str_data) = "" Then
        str1 = str1 & str_data
    End If
Wend

Call CloseFile
read_file = str1

Exit Function
Err:
    read_file = ""
    MsgBox Err.Description
End Function

Private Sub CloseFile()
On Error GoTo Err
Close #1
'CloseFile = True

Exit Sub
Err:
    MsgBox Err.Description
End Sub

Private Sub cmd_request_key_Click()

On Error GoTo Err
    
    If txt_registration_name.Text = "" Then
        MsgBox "Nama Registrasi Masih Kosong...", vbExclamation, headerMSG
        txt_registration_name.SetFocus
        Exit Sub
    End If
    
    If TDBCombo_company.Text = "" Then
        MsgBox "Perusahaan Masih Kosong...", vbExclamation, headerMSG
        TDBCombo_company.SetFocus
        Exit Sub
    End If
    
    If txt_generate_key.Text = "" Then
        MsgBox "File Generate Key Masih Kosong...", vbExclamation, headerMSG
        cmd_req_browse.SetFocus
        Exit Sub
    End If
    
    SQL = "SELECT * FROM s_about WHERE id = '" & h.GetSerialNumber & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        SQL = "UPDATE s_about SET serial = '" & rsa.Encrypt(rsa.PublicKey, h.GetSerialNumber) & "'," & _
                "register = '" & rsa.Encrypt(rsa.PublicKey, txt_registration_name.Text) & "'," & _
                "company = '" & rsa.Encrypt(rsa.PublicKey, TDBCombo_company.Text) & "' " & _
              "WHERE id = '" & h.GetSerialNumber & "'"
        CnG.Execute SQL

    Else
        SQL = "INSERT INTO s_about (id,serial,register,company) " & _
                "VALUES (" & _
                    "'" & h.GetSerialNumber & "','" & rsa.Encrypt(rsa.PublicKey, h.GetSerialNumber) & "','" & rsa.Encrypt(rsa.PublicKey, txt_registration_name.Text) & "'," & _
                    "'" & rsa.Encrypt(rsa.PublicKey, TDBCombo_company.Text) & "')"
        CnG.Execute SQL
    End If
    
    If write_request_key(txt_generate_key.Text) Then
        MsgBox "File Generate Key Berhasil Dibuat...", vbInformation, "CREATE FILE SUCCESS"
    Else
        MsgBox "File Generate Key Gagal Dibuat...", vbCritical, "CREATE FILE ERROR"
    End If
    Exit Sub
          
Err:
MsgBox Err.Description
End Sub

Private Sub load_data_company()
    If rsComp.State Then rsComp.Close
    SQL = "select * from m_company order by company_code"
    rsComp.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsComp
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
    End If
End Sub

Private Sub cmd_req_browse_Click()
    CommonDialog1.Filter = "Key File(*.ssd)|*.ssd"
    CommonDialog1.initDir = App.Path
    CommonDialog1.ShowSave
    
    txt_generate_key.Text = CommonDialog1.FileName
End Sub

Private Sub cmd_reg_browse_Click()
    CommonDialog1.Filter = "Key File(*.ssd)|*.ssd"
    CommonDialog1.initDir = App.Path
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
        txt_register_key.Text = CommonDialog1.FileName
    End If
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub txt_registration_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function write_request_key(ByVal spath As String) As Boolean
On Error GoTo Err
Dim str_data, str1, str2 As String, i As Integer
Dim tglawal As String
    
    Open spath For Output As #1
    Print #1, rsa.Encrypt(rsa.PublicKey, h.GetSerialNumber) & vbCrLf & _
                rsa.Encrypt(rsa.PublicKey, txt_registration_name.Text) & vbCrLf & _
                rsa.Encrypt(rsa.PublicKey, txt_company_name.Text) & vbCrLf & _
                rsa.Encrypt(rsa.PublicKey, h.GetSerialNumber) & vbCrLf & _
                rsa.Encrypt(rsa.PublicKey, Format(Now, "yyyy-MM-dd"))
    Close #1
    
    write_request_key = True
    
    Exit Function
    
Err:
'    MsgBox err.Description
    write_request_key = False
End Function
