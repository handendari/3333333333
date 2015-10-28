VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_mst_history_prestasi 
   Appearance      =   0  'Flat
   Caption         =   "INPUT BONUS"
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
   Icon            =   "frm_mst_history_prestasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   645
      Left            =   12270
      Picture         =   "frm_mst_history_prestasi.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7800
      Width           =   975
   End
   Begin VB.Frame fraAdmin 
      Height          =   7095
      Left            =   240
      TabIndex        =   7
      Top             =   630
      Width           =   12975
      Begin VB.CommandButton cmd_import 
         Caption         =   "&Import"
         Height          =   555
         Left            =   7080
         Picture         =   "frm_mst_history_prestasi.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   6330
         Width           =   855
      End
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
         Height          =   2265
         Left            =   720
         TabIndex        =   10
         Top             =   3870
         Visible         =   0   'False
         Width           =   11535
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   2760
            TabIndex        =   25
            Top             =   240
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   232062979
            CurrentDate     =   40993
         End
         Begin VB.TextBox txt_employee_code 
            Height          =   285
            Left            =   6540
            TabIndex        =   24
            Text            =   "Text3"
            Top             =   240
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.CommandButton cmd_browse_employee 
            Caption         =   "..."
            Height          =   300
            Left            =   4230
            TabIndex        =   23
            Top             =   1020
            Width           =   375
         End
         Begin VB.TextBox txt_employee_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   4680
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   22
            Top             =   1020
            Width           =   2415
         End
         Begin VB.TextBox txt_nik 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2760
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   21
            Top             =   1020
            Width           =   1335
         End
         Begin VB.TextBox txtBonus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   300
            Left            =   2760
            MaxLength       =   40
            TabIndex        =   17
            Top             =   1410
            Width           =   705
         End
         Begin VB.TextBox txt_prestasi_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   300
            Left            =   4080
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   630
            Width           =   3915
         End
         Begin VB.TextBox txtKeterangan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   1
            Top             =   1800
            Width           =   5895
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_prestasi 
            Height          =   375
            Left            =   2760
            OleObjectBlob   =   "frm_mst_history_prestasi.frx":0B20
            TabIndex        =   16
            Top             =   630
            Width           =   1305
         End
         Begin MSAdodcLib.Adodc Adodc_prestasi 
            Height          =   345
            Left            =   3270
            Top             =   690
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tanggal"
            Height          =   195
            Left            =   1560
            TabIndex        =   20
            Top             =   270
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "X GAJI"
            Height          =   195
            Left            =   3540
            TabIndex        =   19
            Top             =   1470
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bonus"
            Height          =   195
            Left            =   1560
            TabIndex        =   18
            Top             =   1440
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Prestasi"
            Height          =   195
            Left            =   1560
            TabIndex        =   14
            Top             =   690
            Width           =   570
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Karyawan"
            Height          =   195
            Left            =   1560
            TabIndex        =   12
            Top             =   1050
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Keterangan"
            Height          =   195
            Left            =   1560
            TabIndex        =   11
            Top             =   1860
            Width           =   840
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
         TabIndex        =   8
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
         Picture         =   "frm_mst_history_prestasi.frx":2ACF
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6330
         Width           =   855
      End
      Begin VB.CommandButton cmdNewUser 
         Caption         =   "&New"
         Height          =   555
         Left            =   720
         Picture         =   "frm_mst_history_prestasi.frx":3059
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6330
         Width           =   855
      End
      Begin VB.CommandButton cmdSimpanUser 
         Caption         =   "&Save"
         Height          =   555
         Left            =   3600
         Picture         =   "frm_mst_history_prestasi.frx":35E3
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6330
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteUser 
         Caption         =   "&Delete"
         Height          =   555
         Left            =   2640
         Picture         =   "frm_mst_history_prestasi.frx":3B6D
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6330
         Width           =   855
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Left            =   1680
         OleObjectBlob   =   "frm_mst_history_prestasi.frx":40F7
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Adodc_company 
         Height          =   345
         Left            =   1680
         Top             =   150
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
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
         TabIndex        =   13
         Top             =   1410
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
         Column(1)       =   "frm_mst_history_prestasi.frx":60B5
         Column(2)       =   "frm_mst_history_prestasi.frx":617D
         FormatStylesCount=   6
         FormatStyle(1)  =   "frm_mst_history_prestasi.frx":6221
         FormatStyle(2)  =   "frm_mst_history_prestasi.frx":6349
         FormatStyle(3)  =   "frm_mst_history_prestasi.frx":63F9
         FormatStyle(4)  =   "frm_mst_history_prestasi.frx":64AD
         FormatStyle(5)  =   "frm_mst_history_prestasi.frx":6585
         FormatStyle(6)  =   "frm_mst_history_prestasi.frx":663D
         ImageCount      =   0
         PrinterProperties=   "frm_mst_history_prestasi.frx":671D
      End
      Begin MSComCtl2.DTPicker DTPicker_monthly 
         Height          =   300
         Left            =   1680
         TabIndex        =   27
         Top             =   870
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   233439235
         UpDown          =   -1  'True
         CurrentDate     =   39278
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   195
         Left            =   720
         TabIndex        =   28
         Top             =   900
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Perusahaan"
         Height          =   195
         Left            =   720
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "INPUT BONUS"
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
      Left            =   5955
      TabIndex        =   26
      Top             =   30
      Width           =   2145
   End
End
Attribute VB_Name = "frm_mst_history_prestasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Private Sub EnableButtonEntryForm _
(ByVal a As Boolean, ByVal b As Boolean, ByVal C As Boolean, ByVal d As Boolean)
'cmdNewForm.Enabled = a And blnUser_Add
'cmdEditForm.Enabled = b And blnUser_Edit
'cmdDeleteForm.Enabled = c And blnUser_Delete
'cmdSaveForm.Enabled = d
End Sub

Private Sub EnableButtonEntryUser _
(ByVal a As Boolean, ByVal b As Boolean, ByVal C As Boolean, ByVal d As Boolean)
cmdNewUser.Enabled = a And blnUser_Add
cmdEditUser.Enabled = b And blnUser_Edit
cmdDeleteUser.Enabled = C And blnUser_Delete
cmdSimpanUser.Enabled = d

'TDBGrid_form.Enabled = Not d
'cmdGenerate.Enabled = Not d
'cmd_save_dtl.Enabled = Not d
End Sub

Private Sub fill_grid_user()
Dim rs1 As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
    Set cmd1.ActiveConnection = CnG
    
    cmd1.CommandText = "select a.performance_date,c.nik,a.prestasi_code,a.kali_gaji,a.description,a.employee_code,c.employee_name " & _
                    "from t_employee_performance a join m_prestasi b on a.prestasi_code = b.kode " & _
                    "join m_employee c on c.employee_code = a.employee_code " & _
                    "where year(performance_date) = '" & Format(DTPicker_monthly.Value, "yyyy") & "' AND " & _
                    "c.company_code = '" & TDBCombo_company.Text & "'"
    rs1.CursorLocation = adUseClient
    rs1.Open cmd1, , adOpenStatic, adLockBatchOptimistic

    With GridEXUser
        Set .ADORecordset = rs1
        .AllowAddNew = False
        .AllowEdit = False
        .AllowDelete = False
        
        With .Columns("performance_date")
            .Caption = "TANGGAL"
            .HeaderAlignment = jgexAlignCenter
            .AllowSizing = False
            .TextAlignment = jgexAlignCenter
            .Width = 1500
        End With
        With .Columns("nik")
            .Caption = "NIK"
            .HeaderAlignment = jgexAlignCenter
            .AllowSizing = False
            .TextAlignment = jgexAlignCenter
            .Width = 1500
        End With
        With .Columns("prestasi_code")
            .Caption = "KODE"
            .HeaderAlignment = jgexAlignCenter
            .AllowSizing = False
            .TextAlignment = jgexAlignCenter
            .Width = 1200
        End With
        With .Columns("kali_gaji")
            .Caption = "KALI GAJI"
            .HeaderAlignment = jgexAlignCenter
            .AllowSizing = False
            .TextAlignment = jgexAlignCenter
            .Width = 1500
        End With
        With .Columns("description")
            .Caption = "DESKRIPSI"
            .HeaderAlignment = jgexAlignCenter
            .AllowSizing = False
            .TextAlignment = jgexAlignLeft
            .Width = 4000
        End With
        With .Columns("employee_code")
            .Caption = "KODE KARY"
            .HeaderAlignment = jgexAlignCenter
            .AllowSizing = False
            .TextAlignment = jgexAlignLeft
            .Width = 1000
            .Visible = False
        End With
        With .Columns("employee_name")
            .Caption = "NAMA KARY"
            .HeaderAlignment = jgexAlignCenter
            .AllowSizing = False
            .TextAlignment = jgexAlignLeft
            .Width = 1000
            .Visible = False
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

Private Sub cmd_browse_employee_Click()
frm_lookup_mst_employee.public_int_mode = 169
frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_import_Click()
frm_trans_import_bonus.Show
End Sub

Private Sub cmdDeleteUser_Click()
If GridEXUser.RowCount < 1 Or GridEXUser.Row < 1 Then Exit Sub

Dim i As Integer
i = MsgBox("Are you sure want to delete data '" _
    & GridEXUser.Value(GridEXUser.Columns("nik").ColPosition) & "' ?", vbOKCancel, headerMSG)

If i = vbOK Then
    CnG.Execute "delete from t_employee_performance where date(performance_date) = '" _
    & Format(GridEXUser.Value(GridEXUser.Columns("performance_date").ColPosition), "yyyy-MM-dd") & "' AND " _
    & "employee_code = '" & GridEXUser.Value(GridEXUser.Columns("employee_code").ColPosition) & "'"
    
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
    
    DTPicker1.Value = Format(GridEXUser.Value(GridEXUser.Columns("performance_date").ColPosition), "yyyy-MM-dd")
    Call set_data_prestasi(GridEXUser.Value(GridEXUser.Columns("prestasi_code").ColPosition))
    txt_employee_code.Text = GridEXUser.Value(GridEXUser.Columns("employee_code").ColPosition)
    txt_nik.Text = GridEXUser.Value(GridEXUser.Columns("nik").ColPosition)
    txt_employee_name.Text = GridEXUser.Value(GridEXUser.Columns("employee_name").ColPosition)
    txtBonus.Text = GridEXUser.Value(GridEXUser.Columns("kali_gaji").ColPosition)
    txtKeterangan.Text = GridEXUser.Value(GridEXUser.Columns("description").ColPosition)

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
    If Not Trim(TDBCombo_prestasi.Text) = "" Then
        CekValidateDataUser = True
    Else
        CekValidateDataUser = False
    End If
End Function


Private Sub cmdNewUser_Click()
If cmdNewUser.Caption = "&New" Then
    cmdNewUser.Caption = "&Cancel"
    Call EnableButtonEntryUser(True, False, False, True)
    Call ShowWindowEntryUser(True)
    fra_EntryUser.Caption = "Entry"
    
    DTPicker1.Value = Now
    TDBCombo_prestasi.Text = ""
    txt_prestasi_name.Text = ""
    txt_employee_code.Text = ""
    txt_nik.Text = ""
    txt_employee_name.Text = ""
    txtBonus.Text = ""
    txtKeterangan.Text = ""
    
    TDBCombo_prestasi.SetFocus
    
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
        MsgBox "Data is not valid", vbCritical, headerMSG
        Exit Sub
    End If
    
    rs.Open "select * from t_employee_performance where prestasi_code ='uOu'", CnG, adOpenKeyset, adLockOptimistic
    
    CnG.BeginTrans
    With rs
        .AddNew
        .Fields("performance_date").Value = Format(DTPicker1.Value, "yyyy-MM-dd")
        .Fields("employee_code").Value = txt_employee_code.Text
        .Fields("prestasi_code").Value = TDBCombo_prestasi.Text
        .Fields("kali_gaji").Value = Val(txtBonus.Text)
        .Fields("description").Value = txtKeterangan.Text
        .Fields("entry_date").Value = Now
        .Fields("entry_user").Value = LOGIN_CODE
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
    
    rs.Open "select * from t_employee_performance where date(performance_date) = '" _
        & Format(GridEXUser.Value(GridEXUser.Columns("performance_date").ColPosition), "yyyy-MM-dd") & "' AND " _
        & "employee_code = '" & GridEXUser.Value(GridEXUser.Columns("employee_code").ColPosition) & "'", CnG, adOpenKeyset, adLockOptimistic
    
    CnG.BeginTrans
    With rs
        
        .Fields("performance_date").Value = Format(DTPicker1.Value, "yyyy-MM-dd")
        .Fields("employee_code").Value = txt_employee_code.Text
        .Fields("prestasi_code").Value = TDBCombo_prestasi.Text
        .Fields("kali_gaji").Value = Val(txtBonus.Text)
        .Fields("description").Value = txtKeterangan.Text
        .Fields("edit_date").Value = Now
        .Fields("edit_user").Value = LOGIN_CODE
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

Private Sub DTPicker_monthly_Change()
Call fill_grid_user
End Sub

Private Sub DTPicker_monthly_Validate(Cancel As Boolean)
Call fill_grid_user
End Sub

Private Sub Form_Load()
Adodc_company.ConnectionString = strConn
Adodc_prestasi.ConnectionString = strConn

Call load_data_company
Call load_data_prestasi
Call load_data_user_access(Me)

Call EnableButtonEntryForm(True, True, True, False)
Call EnableButtonEntryUser(True, True, True, False)
Call ShowWindowEntryUser(False)

DTPicker_monthly.Value = Now

'Call fill_grid_form

Timer1.Enabled = True
End Sub

Public Sub load_data_company()
Adodc_company.RecordSource = "select company_code, company_name from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Public Sub load_data_prestasi()
Adodc_prestasi.RecordSource = "select kode, nama from m_prestasi order by kode"
Adodc_prestasi.Refresh

TDBCombo_prestasi.RowSource = Adodc_prestasi
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

Private Sub TDBCombo_prestasi_ItemChange()
If TDBCombo_prestasi.ApproxCount > 0 Then
    TDBCombo_prestasi.Text = TDBCombo_prestasi.Columns("kode").Value
    txt_prestasi_name = TDBCombo_prestasi.Columns("nama").Value
    
    Call fill_grid_user
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
If LOGIN_LEVEL = 100 Then
    TDBCombo_company.Locked = False
Else
    TDBCombo_company.Locked = True
End If
End Sub

Private Sub set_data_prestasi(ByVal str_code As String)
On Error Resume Next

Adodc_prestasi.Recordset.MoveFirst
Adodc_prestasi.Recordset.Find ("kode = '" & str_code & "'")   ', 0, adSearchForward, 1)
If Not (Adodc_prestasi.Recordset.EOF = True Or Adodc_prestasi.Recordset.BOF = True) Then
    TDBCombo_prestasi.Bookmark = Adodc_prestasi.Recordset.AbsolutePosition
    Call TDBCombo_prestasi_ItemChange
Else
    TDBCombo_prestasi.Text = ""
End If
End Sub
