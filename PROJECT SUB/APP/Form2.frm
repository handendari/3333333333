VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_exec_sql_file 
      Caption         =   "EXEC SQL FILE"
      Height          =   855
      Left            =   6600
      TabIndex        =   12
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox txt_nama 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txt_output 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txt_input 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox txt_key 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Text            =   "HGSDYGDSLWREIUCJD938439402342"
      Top             =   4200
      Width           =   4455
   End
   Begin VB.CommandButton cmd_encrypt 
      Caption         =   "ENCRYPT"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "REPORT"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6120
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn_attendance"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn_attendance"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin CRVIEWERLibCtl.CRViewer CRV 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mFile As CFileSys


Private Sub cmd_encrypt_Click()
txt_output = RC4DeCryptASC(txt_input, txt_key)
End Sub

Public Sub OpenAdoRecordSet()
Dim ts As TextStream
Dim TextFileObj As FileSystemObject
Dim TxtStr As String
Dim rs As New ADODB.Recordset

Set TextFileObj = New FileSystemObject
Set ts = TextFileObj.OpenTextFile("D:\tomcat\projects\attendance\db_backup\sql2.sql", ForReading)
Set rs = New ADODB.Recordset

TxtStr = ts.ReadAll

rs.Open TxtStr, CnG, adOpenForwardOnly, adLockReadOnly


'===========
'Dim oFSO
'Set oFSO = CreateObject("Scripting.FileSystemObject")
'CnG.Execute oFSO.OpenTextFile("D:\tomcat\projects\attendance\db_backup\sql2.sql").ReadAll, 128
''Set CnG = Nothing
'MsgBox "ADO query done"
End Sub

Private Sub cmd_exec_sql_file_Click()
Call OpenAdoRecordSet
End Sub

Private Sub cmd_save_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

cn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;User ID=sa;Data Source=dsn_forwarding"
cn.Open
rs.Open "select * from m_login", cn, adOpenKeyset, adLockOptimistic

With rs
    .AddNew
    !nama = txt_nama
    !Pwd = txt_input
    !pwd_enc = RC4DeCryptASC(txt_input, txt_key)
    .Update
End With
End Sub

Private Sub Command1_Click()
Dim Server As String
Dim DataSourceName As String
Dim DatabaseName As String
Dim Description As String
Dim DriverPath As String
Dim DriverName As String
Dim LastUser As String
Dim Pwd As String

Set mFile = New CFileSys
If mFile.readFile(mFile) Then
    MsgBox "user=" & USER_NAME & vbCr & "pass=" & USER_PASS & vbCr & "host=" & HOST_NAME _
    & vbCr & "db=" & DB_NAME _
    & vbCr & "dsn=" & DSN & vbCr & DESC & vbCr & DRIVER_PATH & vbCr & DRIVER_NAME & vbCr & HD_SN
End If


'MsgBox create_mysql_dsn("localhost", "dsn_TEST___", "root", _
"tomcat123", "db_attendance", , False)
'MsgBox create_mysql_dsn(Server, DataSourceName, LastUser, Pwd, "db_attendance", , False)
End Sub

'Private Function CreateMYSQLDSN _
'    (ByVal svlServerName As String, _
'    ByVal svlDSNName As String, _
'    ByVal svlUserName As String, _
'    ByVal svlPassword As String, _
'    ByVal svlDatabaseName As String, _
'    Optional ByVal svlDSNDriver As String = "MySQL ODBC 5.1 Driver", _
'    Optional ByVal bvlIsUserDSN As Boolean = True) As Boolean
'
'On Error GoTo ErrorHandler
'    CreateMYSQLDSN = False
'    #If Win32 Then
'      Dim intRet As Long
'    #Else
'        Dim intRet As Integer
'    #End If
'
'    strAttributes = "SERVER=" & svlServerName & Chr$(0) & _
'        "DSN=" & svlDSNName & Chr$(0) & _
'        "DESCRIPTION=Created Automatically on " & Format(Date, "DD-MMM-YYYY") & Chr$(0) & _
'        "DATABASE=" & svlDatabaseName & Chr$(0) & _
'        "UID=" & svlUserName & Chr$(0) & _
'        "PWD=" & svlPassword & Chr$(0) & _
'        "PORT=3309" & Chr$(0) & _
'        "OPTION=" & MYSQLOptionValue & Chr$(0)
'    If bvlIsUserDSN Then
'        intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, svlDSNDriver, strAttributes)
'    Else
'        intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, svlDSNDriver, strAttributes)
'    End If
'    If intRet Then
'          CreateMYSQLDSN = True
'          err.Raise vbObjectError + 513, , "DSN " & svlDSNName & " Sucessfully Created..."
'    Else
'        err.Raise vbObjectError + 514, , "Failed to create DSN " & svlDSNName
'    End If
'Exit Function
'
'ErrorHandler:
'    MsgBox ("Error :" & err.Number & " - " & err.Description)
'End Function

Private Sub Command2_Click()
Set mFile = New CFileSys
If mFile.WriteFile(mFile) Then
    MsgBox "OK"
End If
End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "select * from m_user"
Adodc1.Refresh

MsgBox Adodc1.Recordset.Fields(0).Value
End Sub

Private Sub Command4_Click()
Dim a As New frm_rpt

Dim str_sql As String

str_sql = "call spr_attendance_04('2008-07-01','2008-07-31',0,'-',0,'-')"
Call rpt_view(str_sql, "\report\rpt_01.rpt", "TEST PARAMETER")
End Sub

Private Sub Form_Load()
'Call rpt_view

End Sub

Private Sub Form_Resize()

'CRV.Top = 0
'CRV.Left = 0
'CRV.Width = Me.Width - 200
'CRV.Height = Me.Height - 400
End Sub

Public Sub rpt_view _
(ByVal sql_proc As String, ByVal rpt_file As String, ByVal str_param As String)

Dim CrApp As New CRAXDRT.Application
Dim CrRep As New CRAXDRT.Report
Dim AdoRs As New ADODB.Recordset
Dim cr_obj As Object

AdoRs.Open sql_proc, CnG, adOpenDynamic, adLockBatchOptimistic
Set CrRep = CrApp.OpenReport(App.Path & rpt_file)
CrRep.DiscardSavedData
CrRep.Database.Tables(1).SetDataSource AdoRs, 3
crv.ReportSource = CrRep
crv.ViewReport

CrRep.ParameterFields.GetItemByName("p_periode").AddCurrentValue str_param

End Sub
