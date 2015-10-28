VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#5.10#0"; "zkemkeeper.dll"
Begin VB.Form frm_export_user 
   Caption         =   "EXPORT DATA"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   10950
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   615
      Left            =   480
      OleObjectBlob   =   "frm_export_user.frx":0000
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5160
      Top             =   6840
   End
   Begin VB.CommandButton cmdConnect_source 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6720
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton cmd_export 
      Caption         =   "&Upload"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4080
      Picture         =   "frm_export_user.frx":0024
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2880
      Picture         =   "frm_export_user.frx":05AE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   9360
      Picture         =   "frm_export_user.frx":0B38
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "MASTER USER"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "enrollnumber"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "enrollname"
         Caption         =   "NAMA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2594.835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   6000
      Width           =   4695
      _ExtentX        =   8281
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
      Appearance      =   0
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   5535
      Left            =   6720
      TabIndex        =   1
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "DESTINATION (FGP DEVICE)"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "ip_address"
         Caption         =   "IP ADDRESS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "port_number"
         Caption         =   "PORT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1110.047
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6720
      Top             =   6000
      Width           =   3735
      _ExtentX        =   6588
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
      Appearance      =   0
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
      Caption         =   "Adodc2"
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
   Begin VB.Line Line13 
      X1              =   5880
      X2              =   5880
      Y1              =   2280
      Y2              =   2760
   End
   Begin VB.Line Line12 
      X1              =   5880
      X2              =   5880
      Y1              =   3480
      Y2              =   3960
   End
   Begin VB.Line Line11 
      X1              =   5400
      X2              =   5880
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line10 
      X1              =   5400
      X2              =   5880
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line9 
      X1              =   5880
      X2              =   6480
      Y1              =   3960
      Y2              =   3120
   End
   Begin VB.Line Line8 
      X1              =   6480
      X2              =   5880
      Y1              =   3120
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   10440
      Y1              =   6600
      Y2              =   6600
   End
End
Attribute VB_Name = "frm_export_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim glngEnrollData(DATASIZE) As Long
Dim glngEnrollPData As Integer
Dim gbytEnrollData(DATASIZE * 5) As Byte
Dim vEMachineNumber As Integer
Dim jml_rec_upload As Long



Private Sub fill_grid()
Adodc1.RecordSource = "select * from m_enroll order by enrollnumber"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub fill_grid_device()
Adodc2.RecordSource = "select * from m_device order by ip_address"
Adodc2.Refresh

Set DataGrid2.DataSource = Adodc2
End Sub

Private Sub delete_EnrollData()
Dim vEnrollNumber As Integer
Dim vFingerNumber As Integer
Dim vRet As Boolean
Dim vErrorCode As Long

Dim lngReturnCode As Boolean
Dim lngErrorCode As Long
Dim lngCardNumber As Long
Dim par1, par2, par3, par5 As Long
Dim lpszIPAddress As String

'lblMessage.Caption = "Working..."
DoEvents
vEnrollNumber = Val(txtEnrollNumber.Text)
vFingerNumber = cmbBackupNumber.Text

vRet = frmMain.CZKEM1.DeleteEnrollData(frmMain.vMachineNumber, vEnrollNumber, frmMain.vMachineNumber, vFingerNumber)
If vRet = True Then
    'lblMessage.Caption = "DeleteEnrollData OK"
Else
    frmMain.CZKEM1.GetLastError vErrorCode
    'lblMessage.Caption = ErrorPrint(vErrorCode)
End If
    
End Sub

Private Sub cmd_delete_Click()
Dim i As Integer
If Not (DataGrid2.ApproxCount > 0 And DataGrid1.Bookmark > 0) Then Exit Sub

For i = 1 To DataGrid1.ApproxCount
    
Next i
End Sub



Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub parse_enroll_data()
Dim vStr As String
Dim vByte() As Byte
Dim i As Long
 
With Adodc1.Recordset
    If !FingerNumber = 10 Then
        'lstEnrollData.AddItem !Password
        glngEnrollPData = !Password
        
    ElseIf !FingerNumber < 10 Then
        vStr = !FPdata
        vByte = vStr
        For i = 0 To DATASIZE - 1
            glngEnrollData(i) = vByte(i * 5 + 1)
            glngEnrollData(i) = glngEnrollData(i) * 256 + vByte(i * 5 + 2)
            glngEnrollData(i) = glngEnrollData(i) * 256 + vByte(i * 5 + 3)
            glngEnrollData(i) = glngEnrollData(i) * 256 + vByte(i * 5 + 4)
            If vByte(i * 5) = 0 Then
                glngEnrollData(i) = 0 - glngEnrollData(i)
            End If
        Next
    End If
End With
End Sub

Public Sub export_to_device()
Dim vEnrollNumber 'As Integer
Dim vFingerNumber As Integer
Dim vPrivilege As Integer
Dim vRet As Boolean
Dim vErrorCode As Long
Dim i As Integer



With Adodc1.Recordset
    If Not .EOF Then
    
        vEnrollNumber = .Fields("enrollnumber").Value
        vFingerNumber = .Fields("fingernumber").Value
        vPrivilege = .Fields("privilige").Value
        Call parse_enroll_data
        
        vRet = CZKEM1.SetEnrollData(vMachineNumber, vEnrollNumber, vEMachineNumber, _
                vFingerNumber, vPrivilege, glngEnrollData(0), glngEnrollPData)
        If vRet Then
            jml_rec_upload = jml_rec_upload + 1
        End If
    End If
End With
End Sub

Private Sub cmd_export_Click()
If Not DataGrid2.Bookmark > 0 Or Not DataGrid2.ApproxCount > 0 Then
    MsgBox "FGP Device belum terpilih", vbCritical, headerMSG
    Exit Sub
End If

If Trim("" & DataGrid2.Columns("IP ADDRESS").Text) = "" Then
    MsgBox "FGP Device belum terpilih", vbCritical, headerMSG
    Exit Sub
End If

Timer1.Enabled = True
End Sub

Public Sub cmdConnect_source_Click()
Dim bConn As Boolean

If cmdConnect_source.Caption = "Disconnect" Then
    CZKEM1.EnableDevice vMachineNumber, True
    CZKEM1.Disconnect
    cmdConnect_source.Caption = "Connect"

ElseIf cmdConnect_source.Caption = "Connect" Then
    'bConn = CZKEM1.Connect_Net(txtIPAddress_source, CLng(txtPortNo_source))
    bConn = CZKEM1.Connect_Net(DataGrid2.Columns("IP ADDRESS").Value, _
                                DataGrid2.Columns("PORT").Value)
    If bConn Then
        cmdConnect_source.Caption = "Disconnect"
        CZKEM1.EnableDevice vMachineNumber, False
    Else
        'MsgBox "Error connecting to source device...", vbCritical, headerMSG
        Exit Sub
    End If
End If
End Sub

Private Function delete_enroll_data() As Boolean
Dim vEnrollNumber 'As Integer
Dim vFingerNumber As Integer
Dim vRet As Boolean
Dim vErrorCode As Long

Dim lngReturnCode As Boolean
Dim lngErrorCode As Long
Dim lngCardNumber As Long
Dim par1, par2, par3, par5 As Long
Dim lpszIPAddress As String

vEnrollNumber = Val(DataGrid1.Columns("ID").Value)
vFingerNumber = 0

vRet = CZKEM1.DeleteEnrollData(vMachineNumber, vEnrollNumber, vEMachineNumber, vFingerNumber)
If vRet Then
    'MsgBox "Data " & vEnrollNumber & " has been deleted", vbInformation, headerMSG
    delete_enroll_data = True
Else
    'CZKEM1.GetLastError vErrorCode
    'MsgBox ErrorPrint(vErrorCode)
    delete_enroll_data = False
End If

vRet = CZKEM1.RefreshData(vMachineNumber)
End Function

Private Sub cmdDelete_Click()
Dim i As Integer
Dim j As Boolean, str_msg As String

i = MsgBox("Anda yakin menghapus data ID '" & DataGrid1.Columns("ID").Value & "'" & vbCr _
& "dari device " & DataGrid2.Columns("IP ADDRESS").Value & " ?", vbOKCancel, headerMSG)

If i = vbCancel Then
    Exit Sub
End If

If cmdConnect_source.Caption = "Connect" Then Call cmdConnect_source_Click
If Not cmdConnect_source.Caption = "Disconnect" Then
    MsgBox "Error connecting to destination device...", vbCritical, headerMSG
    Exit Sub
End If

j = delete_enroll_data

If cmdConnect_source.Caption = "Disconnect" Then Call cmdConnect_source_Click
If Not cmdConnect_source.Caption = "Connect" Then
    MsgBox "Error disconnecting from destination device...", vbCritical, headerMSG
    Exit Sub
End If

If j = True Then
    str_msg = "Data '" & DataGrid1.Columns("ID").Value & "' has been deleted"
Else
    str_msg = "No Data has been deleted"
End If
MsgBox str_msg, vbInformation, headerMSG
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Adodc1.Caption = "Record " & Adodc1.Recordset.AbsolutePosition & " of " & Adodc1.Recordset.RecordCount
End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Adodc2.Caption = "Record " & Adodc2.Recordset.AbsolutePosition & " of " & Adodc2.Recordset.RecordCount
End Sub

Private Sub Form_Load()
Me.Width = 11070
Me.Height = 8445

Adodc1.ConnectionString = strConn
Adodc2.ConnectionString = strConn

Call fill_grid
Call fill_grid_device

DataGrid1.Columns("NAMA").Visible = False

Call DataGrid1_RowColChange(1, 1)
Call DataGrid2_RowColChange(1, 1)

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False

If cmdConnect_source.Caption = "Connect" Then Call cmdConnect_source_Click
If Not cmdConnect_source.Caption = "Disconnect" Then
    MsgBox "Error connecting to destination device...", vbCritical, headerMSG
    Exit Sub
End If

If Adodc1.Recordset.RecordCount > 0 Then Adodc1.Recordset.MoveFirst

jml_rec_upload = 0
frm_progess1.Caption = "Uploading..."
frm_progess1.Show 1

If cmdConnect_source.Caption = "Disconnect" Then Call cmdConnect_source_Click
If Not cmdConnect_source.Caption = "Connect" Then
    MsgBox "Error disconnecting from destination device...", vbCritical, headerMSG
    Exit Sub
End If

MsgBox jml_rec_upload & " data have exported successfully...", vbInformation, headerMSG
If Adodc1.Recordset.RecordCount > 0 Then Adodc1.Recordset.MoveFirst
End Sub
