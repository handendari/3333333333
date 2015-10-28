VERSION 5.00
Begin VB.Form frm_change_password 
   Caption         =   "UBAH PASSWORD"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6255
   Icon            =   "frm_change_password.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3330
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmPerubahanPasswd 
      Caption         =   "Ubah Password"
      Height          =   2985
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   6015
      Begin VB.TextBox txt_konfirm_pass 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2250
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1440
         Width           =   3405
      End
      Begin VB.TextBox txt_new_pass 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2250
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   3405
      End
      Begin VB.TextBox txt_old_pass 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2250
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   720
         Width           =   3405
      End
      Begin prj_panji.vbButton cmdGantiPasswd 
         Height          =   450
         Left            =   2850
         TabIndex        =   3
         Top             =   2190
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   794
         BTYPE           =   14
         TX              =   "&Ubah"
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
         MICON           =   "frm_change_password.frx":1082
         PICN            =   "frm_change_password.frx":109E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdBatal 
         Height          =   450
         Left            =   4200
         TabIndex        =   4
         Top             =   2190
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
         MICON           =   "frm_change_password.frx":2130
         PICN            =   "frm_change_password.frx":214C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Konfirmasi Password"
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Top             =   1500
         Width           =   2085
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Password Baru"
         Height          =   285
         Left            =   570
         TabIndex        =   7
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Password Lama"
         Height          =   285
         Left            =   570
         TabIndex        =   6
         Top             =   780
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_change_password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vNewPwd As String

Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdGantiPasswd_Click()
Dim strsql As String
Dim rs As New ADODB.Recordset
Dim v_password As String

'+++++++++++++++ Cek Password Lama +++++++++++++++++++++++
strsql = "SELECT user_pass FROM m_user WHERE user_name = '" & LOGIN_NAME & "'"
rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

If rs.RecordCount > 0 Then
    v_password = DecryptINI(Trim(rs!USER_PASS), pEncryptionPassword)
End If

If txt_old_pass <> v_password Then
    MsgBox "Password Lama Tidak Sesuai!" & Chr(13) & _
        "Silahkan Coba Lagi..", vbExclamation, headerMSG
    txt_old_pass.Text = ""
    txt_old_pass.SetFocus
    Exit Sub
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'+++++++++++++++ Cek Password Baru +++++++++++++++++++++++
If txt_new_pass <> txt_konfirm_pass Then
    MsgBox "Password Baru Tidak Sesuai!" & Chr(13) & _
        "Silahkan Coba Lagi..", vbExclamation, headerMSG
    txt_konfirm_pass.Text = ""
    txt_konfirm_pass.SetFocus
    Exit Sub
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++

If LOGIN_LEVEL = 100 Then
    strsql = "UPDATE m_user SET user_pass = '" & EncryptINI(Trim(txt_new_pass.Text), pEncryptionPassword) & "' " & _
            "WHERE user_name = '" & LOGIN_NAME & "'"
Else
    strsql = "UPDATE m_user SET user_pass = '" & EncryptINI(Trim(txt_new_pass.Text), pEncryptionPassword) & "'," & _
                "flag_ubah_pwd = 1 " & _
             "WHERE user_name = '" & LOGIN_NAME & "'"
End If
CnG.Execute strsql

MsgBox "Ubah Password Sukses!", vbInformation, headerMSG

vNewPwd = txt_new_pass.Text

txt_konfirm_pass.Text = ""
txt_new_pass.Text = ""
txt_old_pass.Text = ""

load_login
        
Unload Me
End Sub

Private Sub txt_konfirm_pass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdGantiPasswd_Click
End Sub

Private Sub load_login()
Dim rs As New ADODB.Recordset
Dim str_sql As String
Dim v_company As String

    str_sql = "select a.user_code,a.user_name,a.full_name,a.user_pass,a.company_code," & _
                    "a.flag_company_access,b.level_code, a.user_level,now() tglserver,b.division_code department_code," & _
                    "a.flag_user,b.employee_name,a.flag_pwd_awal,a.flag_ubah_pwd " & _
            "from m_user a left join m_employee b on a.employee_code = b.employee_code " & _
            "where user_name = '" & Replace(LOGIN_NAME, "'", "''") & "' " & _
            "and user_pass = '" & EncryptINI(vNewPwd, pEncryptionPassword) & "'"
            
    rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        loginTime = rs!tglServer
        LOGIN_CODE = rs!user_code
        LOGIN_NAME = rs!USER_NAME
        LOGIN_FULLNAME = IIf(IsNull(rs!full_name), "", rs!full_name)
        LOGIN_PASS = rs!USER_PASS
        LOGIN_LEVEL = rs!user_level
        DEPARTMENT_CODE = IIf(IsNull(rs!DEPARTMENT_CODE), "", rs!DEPARTMENT_CODE)
        EMPLOYEE_NAME = IIf(IsNull(rs!EMPLOYEE_NAME), "", rs!EMPLOYEE_NAME)
        vFlagChgPwd = IIf(IsNull(rs!flag_pwd_awal), 0, rs!flag_pwd_awal)
        vFlagUbahPwd = IIf(IsNull(rs!flag_ubah_pwd), 0, rs!flag_ubah_pwd)
        
    '    If LOGIN_LEVEL = 100 Then
    '        DATA_LEVEL = 1
    '    Else
    '        DATA_LEVEL = IIf(IsNull(rs!level_code), 10000, rs!level_code)
    '    End If
        
        COMPANY_CODE = rs!COMPANY_CODE
        COMPANY_ACCESS = IIf(IsNull(rs!flag_user) = True, 0, rs!flag_user)
        
        If rs.RecordCount = 1 Then
            If Not Trim(rs!USER_NAME) = Trim(LOGIN_NAME) Or _
                Not Trim(rs!USER_PASS) = EncryptINI(Trim(vNewPwd), pEncryptionPassword) Then
                MsgBox "Username / Password Tidak Sesuai...", vbCritical, headerMSG
                intOKTime = intOKTime + 1
                If intOKTime >= 3 Then End
                Exit Sub
            End If
        End If
    rs.Close
    
        '++++++++++INSERT UNTUK LOG AKTIFITAS USER++++++++++++
        Dim clsFn As New clsFunction
        clsFn.InsertLog ("Ubah Password")
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++
    Else
        MsgBox "Username / Password Tidak Sesuai...", vbCritical, headerMSG
        intOKTime = intOKTime + 1
        If intOKTime >= 3 Then End
        Exit Sub
    End If
        
    If BLN_RUNNING = True Then
        Unload Me
        
        If vFlagChgPwd = 1 And vFlagUbahPwd = 0 Then
            Call frm_etc_login.close_all_forms
            Call mdi_absensi.re_create_menu
            
            frm_change_password.Show 1
        Else
            Call frm_etc_login.close_all_forms
            Call mdi_absensi.re_create_menu
            'Call mdi_absensi.reload_user_privileges
        End If
    Else
        Call mdi_absensi.re_create_menu
        Unload Me
        
        CnG.Execute "call spg_leave_periode2 ('" & Format(Now, "yyyy-MM-dd") & "')"
        CnG.Execute "DELETE from temp_list"
        
        If vFlagChgPwd = 1 And vFlagUbahPwd = 0 Then
            mdi_absensi.Show
            
            frm_change_password.Show 1
        Else
            mdi_absensi.Show
        End If
    End If
End Sub
