VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_mst_prestasi 
   Appearance      =   0  'Flat
   Caption         =   "MASTER PRESTASI KARYAWAN"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_mst_prestasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   645
      Left            =   12180
      Picture         =   "frm_mst_prestasi.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7230
      Width           =   975
   End
   Begin VB.Frame fraAdmin 
      Height          =   6855
      Left            =   240
      TabIndex        =   9
      Top             =   990
      Width           =   12975
      Begin VB.Frame fra_EntryUser 
         Caption         =   "Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   720
         TabIndex        =   12
         Top             =   4080
         Visible         =   0   'False
         Width           =   11535
         Begin VB.TextBox txt_user_code 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   300
            Left            =   2760
            MaxLength       =   40
            TabIndex        =   1
            Top             =   300
            Width           =   1335
         End
         Begin VB.TextBox txt_NamaUser 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   300
            Left            =   2760
            MaxLength       =   40
            TabIndex        =   2
            Top             =   660
            Width           =   2895
         End
         Begin VB.TextBox txt_PasswordUser 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   3
            Top             =   1020
            Width           =   5895
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "KODE"
            Height          =   195
            Left            =   1560
            TabIndex        =   16
            Top             =   300
            Width           =   405
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "NAMA"
            Height          =   195
            Left            =   1560
            TabIndex        =   14
            Top             =   660
            Width           =   435
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "DESKRIPSI"
            Height          =   195
            Left            =   1560
            TabIndex        =   13
            Top             =   1020
            Width           =   780
         End
      End
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   300
         Left            =   3360
         Locked          =   -1  'True
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   4815
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   120
         Top             =   5880
      End
      Begin VB.CommandButton cmdEditUser 
         Caption         =   "&Edit"
         Height          =   555
         Left            =   1680
         Picture         =   "frm_mst_prestasi.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6000
         Width           =   855
      End
      Begin VB.CommandButton cmdNewUser 
         Caption         =   "&New"
         Height          =   555
         Left            =   720
         Picture         =   "frm_mst_prestasi.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6000
         Width           =   855
      End
      Begin VB.CommandButton cmdSimpanUser 
         Caption         =   "&Save"
         Height          =   555
         Left            =   3600
         Picture         =   "frm_mst_prestasi.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6000
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteUser 
         Caption         =   "&Delete"
         Height          =   555
         Left            =   2640
         Picture         =   "frm_mst_prestasi.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6000
         Width           =   855
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Left            =   1680
         OleObjectBlob   =   "frm_mst_prestasi.frx":1BBE
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Adodc_company 
         Height          =   375
         Left            =   1680
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEXUser 
         Height          =   4725
         Left            =   720
         TabIndex        =   15
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   8334
         Version         =   "2.0"
         HeaderStyle     =   3
         MethodHoldFields=   -1  'True
         ForeColorHeader =   -2147483639
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         BackColorHeader =   -2147483646
         HeaderFontName  =   "Microsoft Sans Serif"
         HeaderFontSize  =   9
         ColumnHeaderHeight=   315
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frm_mst_prestasi.frx":3B7C
         Column(2)       =   "frm_mst_prestasi.frx":3C44
         FormatStylesCount=   6
         FormatStyle(1)  =   "frm_mst_prestasi.frx":3CE8
         FormatStyle(2)  =   "frm_mst_prestasi.frx":3E10
         FormatStyle(3)  =   "frm_mst_prestasi.frx":3EC0
         FormatStyle(4)  =   "frm_mst_prestasi.frx":3F74
         FormatStyle(5)  =   "frm_mst_prestasi.frx":404C
         FormatStyle(6)  =   "frm_mst_prestasi.frx":4104
         ImageCount      =   0
         PrinterProperties=   "frm_mst_prestasi.frx":41E4
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Perusahaan"
         Height          =   195
         Left            =   720
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "MASTER PRESTASI"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6150
      TabIndex        =   17
      Top             =   0
      Width           =   2955
   End
End
Attribute VB_Name = "frm_mst_prestasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Private Sub EnableButtonEntryForm _
(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, ByVal d As Boolean)
'cmdNewForm.Enabled = a And blnUser_Add
'cmdEditForm.Enabled = b And blnUser_Edit
'cmdDeleteForm.Enabled = c And blnUser_Delete
'cmdSaveForm.Enabled = d
End Sub

Private Sub EnableButtonEntryUser _
(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, ByVal d As Boolean)
cmdNewUser.Enabled = a And blnUser_Add
cmdEditUser.Enabled = b And blnUser_Edit
cmdDeleteUser.Enabled = c And blnUser_Delete
cmdSimpanUser.Enabled = d

'TDBGrid_form.Enabled = Not d
'cmdGenerate.Enabled = Not d
'cmd_save_dtl.Enabled = Not d
End Sub

Private Sub fill_grid_user()
Dim rs1 As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
    Set cmd1.ActiveConnection = CnG
    
    cmd1.CommandText = "select a.kode,a.nama,a.ket " & _
                    "from m_prestasi a order by kode"
    rs1.CursorLocation = adUseClient
    rs1.Open cmd1, , adOpenStatic, adLockBatchOptimistic

    With GridEXUser
        Set .ADORecordset = rs1
        .AllowAddNew = False
        .AllowEdit = False
        .AllowDelete = False
        
        With .Columns("kode")
            .Caption = "CODE"
            .HeaderAlignment = jgexAlignCenter
            .AllowSizing = False
            .TextAlignment = jgexAlignLeft
            .Width = 2000
        End With
        With .Columns("nama")
            .Caption = "NAME"
            .HeaderAlignment = jgexAlignCenter
            .AllowSizing = False
            .TextAlignment = jgexAlignLeft
            .Width = 3000
        End With
        With .Columns("ket")
            .Caption = "REMARK"
            .HeaderAlignment = jgexAlignCenter
            .AllowSizing = False
            .TextAlignment = jgexAlignLeft
            .Width = 3000
        End With
    End With
End Sub

Private Sub ShowWindowEntryUser(ByVal i As Boolean)
If i = True Then ' Max
    fra_EntryUser.Visible = True
    'GridEXUser.Height = 960
ElseIf i = False Then ' Min
    fra_EntryUser.Visible = False
    'GridEXUser.Height = 1800
End If
End Sub

Private Sub cmdDeleteUser_Click()
If GridEXUser.RowCount < 1 Or GridEXUser.Row < 1 Then Exit Sub

Dim i As Integer
i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
    & GridEXUser.Value(GridEXUser.Columns("user_name").ColPosition) & "' ?", vbOKCancel, headerMSG)

If i = vbOK Then
    CnG.Execute "delete from m_prestasi where kode = '" _
    & GridEXUser.Value(GridEXUser.Columns("kode").ColPosition) & "'"
    
    Call fill_grid_user
End If
End Sub

Private Sub cmdEditUser_Click()
If GridEXUser.RowCount < 1 Or GridEXUser.Row < 1 Then Exit Sub

If cmdEditUser.Caption = "&Edit" Then
    cmdEditUser.Caption = "&Cancel"
    Call EnableButtonEntryUser(False, True, False, True)
    Call ShowWindowEntryUser(True)
    fra_EntryUser.Caption = "Edit"
    
    txt_user_code = GridEXUser.Value(GridEXUser.Columns("kode").ColPosition)
    txt_NamaUser = GridEXUser.Value(GridEXUser.Columns("nama").ColPosition)
    txt_PasswordUser = GridEXUser.Value(GridEXUser.Columns("ket").ColPosition)
    
    txt_NamaUser.SelStart = 0
    txt_NamaUser.SelLength = Len(Trim(txt_NamaUser))
    txt_NamaUser.SetFocus
    Call EnabledOptionUser(False)
Else
    cmdEditUser.Caption = "&Edit"
    Call EnableButtonEntryUser(True, True, True, False)
    Call ShowWindowEntryUser(False)
    Call EnabledOptionUser(True)
End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Function CekValidateDataUser() As Boolean
    If Not Trim(txt_NamaUser) = "" Then
        CekValidateDataUser = True
    Else
        CekValidateDataUser = False
    End If
End Function

Private Function CekDuplicateNameUser() As Boolean
Dim cmd1 As New ADODB.Command
Dim rs1 As New ADODB.Recordset
    Set cmd1.ActiveConnection = CnG
    cmd1.CommandText = "select 1 from m_prestasi " & _
                       "where kode = '" & txt_user_code.Text & "'"
    rs1.CursorLocation = adUseClient
    rs1.Open cmd1, , adOpenStatic, adLockBatchOptimistic
    
    If rs1.RecordCount > 0 Then
        CekDuplicateNameUser = True
    Else
        CekDuplicateNameUser = False
    End If
End Function

Private Sub cmdNewUser_Click()
If cmdNewUser.Caption = "&New" Then
    cmdNewUser.Caption = "&Cancel"
    Call EnableButtonEntryUser(True, False, False, True)
    Call ShowWindowEntryUser(True)
    fra_EntryUser.Caption = "Entry"
    
    txt_user_code = ""
    txt_NamaUser = ""
    txt_PasswordUser = ""
    
    txt_user_code.SetFocus
    
    Call EnabledOptionUser(False)
Else
    cmdNewUser.Caption = "&New"
    Call EnableButtonEntryUser(True, True, True, False)
    Call ShowWindowEntryUser(False)
    Call EnabledOptionUser(True)
End If
End Sub

Private Sub cmdSimpanUser_Click()
Dim rs As New ADODB.Recordset
Dim int_level As Integer
Dim clsFunc As New clsFunction

If fra_EntryUser.Caption = "Entry" Then
    If CekValidateDataUser = False Then
        MsgBox "Data Tidak Valid...", vbCritical, headerMSG
        Exit Sub
    End If
    If CekDuplicateNameUser = True Then
        MsgBox "This Data is Already Exist...", vbCritical, headerMSG
        Exit Sub
    End If
    
    rs.Open "select * from m_prestasi where kode='uOu'", CnG, adOpenKeyset, adLockOptimistic
    
    CnG.BeginTrans
    With rs
        .AddNew
        
        .Fields("kode").Value = Trim(txt_user_code)
        .Fields("nama").Value = Trim(txt_NamaUser)
        .Fields("ket").Value = txt_PasswordUser.Text
        .Fields("user_input").Value = LOGIN_CODE
        .Fields("entry_date").Value = Now
        .Update
    End With
    CnG.CommitTrans
    
    Call fill_grid_user
    Call EnableButtonEntryUser(True, True, True, False)
    Call ShowWindowEntryUser(False)
    cmdNewUser.Caption = "&New"
    Call EnabledOptionUser(True)
    
ElseIf fra_EntryUser.Caption = "Edit" Then
    If CekValidateDataUser = False Then
        MsgBox "Editing Data Not Valid", vbCritical, "Request validate data"
        Exit Sub
    End If
    
    rs.Open "select * from m_prestasi where kode='" _
    & GridEXUser.Value(GridEXUser.Columns("kode").ColPosition) & "'", CnG, adOpenKeyset, adLockOptimistic
    
    CnG.BeginTrans
    With rs
        
        .Fields("kode").Value = Trim(txt_user_code)
        .Fields("nama").Value = Trim(txt_NamaUser)
        .Fields("ket").Value = txt_PasswordUser.Text
        .Fields("user_edit").Value = LOGIN_CODE
        .Fields("edit_date").Value = Now
        .Update
    End With
    CnG.CommitTrans
    
    Call fill_grid_user
    Call EnableButtonEntryUser(True, True, True, False)
    Call ShowWindowEntryUser(False)
    cmdEditUser.Caption = "&Edit"
    Call EnabledOptionUser(True)
End If
End Sub

Private Sub Form_Load()
Adodc_company.ConnectionString = strConn

Call load_data_company
Call load_data_user_access(Me)

Call EnableButtonEntryForm(True, True, True, False)
Call EnableButtonEntryUser(True, True, True, False)
Call ShowWindowEntryUser(False)

'Call fill_grid_form

timer1.Enabled = True
End Sub

Public Sub load_data_company()
Adodc_company.RecordSource = "select company_code, company_name from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub EnabledOptionUser(ByVal i As Boolean)
'fraOption.Enabled = i
GridEXUser.Enabled = i
End Sub

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
    
    Call fill_grid_user
End If
End Sub

Private Sub Timer1_Timer()
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
If LOGIN_LEVEL = 100 Then
    TDBCombo_company.Locked = False
Else
    TDBCombo_company.Locked = True
End If
End Sub
