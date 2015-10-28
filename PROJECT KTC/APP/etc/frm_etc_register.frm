VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_etc_register 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "REGISTER THIS PRODUCT"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7065
   Icon            =   "frm_etc_register.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_load 
      Caption         =   "&Load Key"
      Height          =   645
      Left            =   2400
      Picture         =   "frm_etc_register.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Height          =   315
      Left            =   480
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmd_request_key 
      Caption         =   "Req. Key"
      Height          =   645
      Left            =   1320
      Picture         =   "frm_etc_register.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   6135
      Begin VB.TextBox txt_license_key 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   30
         PasswordChar    =   "#"
         TabIndex        =   7
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txt_finger_print 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   6
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "LICENSE KEY"
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FINGER PRINT"
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Width           =   1140
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   645
      Left            =   5640
      Picture         =   "frm_etc_register.frx":0B20
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmd_register 
      Caption         =   "&Register"
      Height          =   645
      Left            =   3480
      Picture         =   "frm_etc_register.frx":10AA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   975
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
      TabIndex        =   2
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   360
      Picture         =   "frm_etc_register.frx":1634
      Stretch         =   -1  'True
      Top             =   360
      Width           =   6240
   End
End
Attribute VB_Name = "frm_etc_register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmd_register_Click()
If txt_license_key.Text = RC4DeCryptASC(txt_finger_print.Text, pEncryptionPassword) Then
    Call set_license
    Call Form_Load
    If Check1.Value = vbChecked Then
        MsgBox "This Product was licensed successfully!", vbInformation, headerMSG
        IS_LICENSED = True
        If MDI_STATE = True Then
            Call mdi_absensi.show_app_title
        Else
            Unload Me
            frm_etc_login.timer1.Enabled = True
            frm_etc_login.Show
        End If
    End If
Else
    MsgBox "The license key is not match!", vbCritical, headerMSG
End If
End Sub

Private Sub cmd_request_key_Click()
Dim zip_file As String, Ret
Dim rs1 As New ADODB.Recordset
Dim str1, str2 As String

rs1.Open "select * from m_company order by company_code asc limit 0,1", CnG, adOpenStatic, adLockReadOnly
If rs1.RecordCount > 0 Then

    str2 = "                  : "
'    str1 = get_add_str("COMPANY") & rs1.Fields("company_name").Value & vbCr _
'            & get_add_str("ADDRESS") & rs1.Fields("address").Value & vbCr _
'            & get_add_str("POSTCODE") & rs1.Fields("postal_code").Value & vbCr _
'            & get_add_str("CITY") & rs1.Fields("city_name").Value & vbCr _
'            & get_add_str("PHONE") & rs1.Fields("phone_number").Value & vbCr _
'            & get_add_str("FAX") & rs1.Fields("fax_number").Value
    str1 = get_add_str("COMPANY") & rs1.Fields("company_name").Value _
            & Chr(13) & get_add_str("ADDRESS") & rs1.Fields("address").Value _
            & Chr(13) & get_add_str("POSTCODE") & rs1.Fields("postal_code").Value _
            & Chr(13) & get_add_str("CITY") & rs1.Fields("city_name").Value _
            & Chr(13) & get_add_str("PHONE") & rs1.Fields("phone_number").Value _
            & Chr(13) & get_add_str("FAX") & rs1.Fields("fax_number").Value
End If

'zip_file = App.Path & "\test.zip"
ShellExecute Me.hWnd, "Open", "mailto:register@solusisentraldata.com?" _
    & "subject=" & "Request License Key for Attendance and Payroll Software (" & Format(Now, "dd-mm-yyyy") & ")" _
    & "&body=" & get_add_str("FINGER PRINT") & Trim(txt_finger_print) & vbCr & str1, vbNullString, vbNullString, vbNormalFocus
DoEvents
'While Ret = 0
'    Ret = FindWindow(vbNullString, zip_file)
'Wend
'SendKeys ("%ia" & "" & "{TAB}{TAB}{ENTER}")
'SendKeys ("{TAB}{TAB}{ENTER}")
'To send automatically change the
'ShellExecute vbNormalFocus to vbHide and add:
'SendKeys ("%s")
'MsgBox "Message with attachment sent."


'Dim rs1 As New ADODB.Recordset
'Dim str1, str2 As String
'Dim oItem As Object
'
'Dim oLapp As New Outlook.Application
'Set oItem = oLapp.CreateItem(0)
'
'rs1.Open "select * from m_company order by company_code asc limit 0,1", CnG, adOpenStatic, adLockReadOnly
'If rs1.RecordCount > 0 Then
'
'    str2 = "                  : "
'    str1 = get_add_str("COMPANY") & rs1.Fields("company_name").Value & vbCr _
'            & get_add_str("ADDRESS") & rs1.Fields("address").Value & vbCr _
'            & get_add_str("POSTCODE") & rs1.Fields("postal_code").Value & vbCr _
'            & get_add_str("CITY") & rs1.Fields("city_name").Value & vbCr _
'            & get_add_str("PHONE") & rs1.Fields("phone_number").Value & vbCr _
'            & get_add_str("FAX") & rs1.Fields("fax_number").Value
'End If
'
'
'With oItem
'   .Subject = "Request License Key for Payroll Application (" & Format(Now, "dd-mm-yyyy") & ")"
'   .To = "tabrani@solusisentraldata.com"
'   .Body = get_add_str("FINGER PRINT") & Trim(txt_finger_print) & vbCr & str1
'   .Display
'End With
End Sub

Private Function get_add_str(ByVal str1 As String) As String
Dim str2 As String

str2 = "                  : "
get_add_str = str1 & Right(str2, 15 - Len(str1))
End Function

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Function get_finger_print() As String
Dim h As New HDSN
Dim str1 As String
Dim rs1 As New ADODB.Recordset

get_finger_print = h.GetSerialNumber
'(h.GetModelNumber & "" & h.GetSerialNumber)
End Function

'Private Function check_license(ByVal str_id) As Boolean
'Dim h As New HDSN
'Dim str1 As String
'Dim rs1 As New ADODB.Recordset
'
'rs1.Open "select * from m_license order by drive_id asc", CnG, adOpenStatic, adLockReadOnly
'
'While Not rs1.EOF
'
'    h.CurrentDrive = Val(rs1.Fields("drive_id").Value)
'    If rs1.Fields("license_number").Value = RC4DeCryptASC(str_id, pEncryptionPassword) Then
'        check_license = True
'        Exit Function
'    Else
'        check_license = False
'    End If
'
'    rs1.MoveNext
'
'Wend
'End Function

Private Sub Form_Load()
txt_finger_print.Text = get_finger_print
Check1.Value = IIf(check_license = True, vbChecked, vbUnchecked)

If Check1.Value = vbChecked Then
    lbl_license.Caption = "This Program has licensed"
    txt_license_key = get_license_key
    
    cmd_request_key.Enabled = False
    cmd_load.Enabled = False
    cmd_register.Enabled = False
Else
    lbl_license.Caption = "This Program has no license"
    txt_license_key = ""
    
    cmd_request_key.Enabled = True
    cmd_load.Enabled = True
    cmd_register.Enabled = True
End If
End Sub

Private Function get_license_key() As String
Dim h As New HDSN, str_l As String
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
On Error GoTo err

If Mode = 0 Then Open str_filename For Input As #1
If Mode = 1 Then Open str_filename For Output As #1
OpenFile = True

Exit Function
err:
    OpenFile = False
    MsgBox err.Description
End Function

Private Sub cmd_load_Click()
CommonDialog1.Filter = "ssd|*.ssd"
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowOpen

If CommonDialog1.FileName <> "" Then
    txt_license_key = read_file(CommonDialog1.FileName)
End If
End Sub

Private Function read_file(ByVal spath As String) As String
Dim mFile1 As New CFileSys

On Error GoTo err

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
err:
    read_file = ""
    MsgBox err.Description
End Function

Private Sub CloseFile()
On Error GoTo err
Close #1
'CloseFile = True

Exit Sub
err:
    MsgBox err.Description
End Sub

