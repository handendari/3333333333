VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_master_enroll 
   Caption         =   "EXPORT DATA"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   10950
   Begin VB.CommandButton cmd_delete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   7440
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6255
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   11033
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      Top             =   6720
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
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "frm_master_enroll.frx":0000
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   6255
      Left            =   6720
      TabIndex        =   4
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   11033
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
      Caption         =   "FGP DEVICE"
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
      Top             =   6720
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
   Begin VB.Line Line7 
      X1              =   6000
      X2              =   6000
      Y1              =   2400
      Y2              =   3120
   End
   Begin VB.Line Line6 
      X1              =   6000
      X2              =   6000
      Y1              =   3840
      Y2              =   4560
   End
   Begin VB.Line Line5 
      X1              =   5280
      X2              =   6000
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line4 
      X1              =   5280
      X2              =   6000
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   6000
      X2              =   6600
      Y1              =   4560
      Y2              =   3480
   End
   Begin VB.Line Line2 
      X1              =   6600
      X2              =   6000
      Y1              =   3480
      Y2              =   2400
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   10440
      Y1              =   7200
      Y2              =   7200
   End
End
Attribute VB_Name = "frm_master_enroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim glngEnrollData(DATASIZE) As Long
Dim glngEnrollPData As Integer
Dim gbytEnrollData(DATASIZE * 5) As Byte
Dim vEMachineNumber As Integer



Private Sub get_data_source()

Dim vEnrollNumber 'As Integer
Dim vFingerNumber As Integer
Dim vPrivilege As Integer
Dim vEnable As Integer
Dim vFlag As Boolean
Dim vRet As Boolean
Dim vErrorCode As Long
Dim vStr As String
Dim i As Long
Dim j
Dim rs As New ADODB.Recordset


vRet = CZKEM1.ReadAllUserID(vMachineNumber)
If vRet Then
    'lblMessage.Caption = "ReadAllUserID OK"
Else
    CZKEM1.GetLastError vErrorCode
    MsgBox ErrorPrint(vErrorCode), vbCritical, headerMSG
    Exit Sub
End If

'---- Get Source Enroll data and save into database -------------
MousePointer = vbHourglass
vFlag = False

Adodc1.RecordSource = "select * from tblenroll order by enrollnumber asc"
Adodc1.Refresh

j = lng_start
Do While j >= lng_start And j <= lng_end
        
    vFlag = True
    vEnrollNumber = j
        
    vRet = CZKEM1.GetEnrollData(vMachineNumber, vEnrollNumber, vEMachineNumber, _
            vFingerNumber, vPrivilege, glngEnrollData(0), glngEnrollPData)
       
    If vRet = True Then
    
        If rs.State = 1 Then rs.Close
        rs.Open "select count(*) as jml from tblEnroll where enrollnumber = " & vEnrollNumber, Cn, adOpenStatic, adLockReadOnly
        
        If rs.Fields("jml").Value > 0 Then
            Cn.Execute "delete tblenroll where enrollnumber = " & vEnrollNumber
        End If
        
        With Adodc1.Recordset
        .AddNew
        !EMachineNumber = vMachineNumber
        !EnrollNumber = vEnrollNumber
        !FingerNumber = vFingerNumber
        !Privilige = vPrivilege
        
        If vFingerNumber = 10 Then
            !Password = glngEnrollPData
        Else
            For i = 0 To DATASIZE - 1
                gbytEnrollData(i * 5) = 1
                If glngEnrollData(i) < 0 Then
                    gbytEnrollData(i * 5) = 0
                    glngEnrollData(i) = Abs(glngEnrollData(i))
                End If
                gbytEnrollData(i * 5 + 1) = (glngEnrollData(i) \ 256 \ 256 \ 256)
                gbytEnrollData(i * 5 + 2) = (glngEnrollData(i) \ 256 \ 256) Mod 256
                gbytEnrollData(i * 5 + 3) = (glngEnrollData(i) \ 256) Mod 256
                gbytEnrollData(i * 5 + 4) = glngEnrollData(i) Mod 256
            Next
            !FPdata = gbytEnrollData
        End If
        .Update
        End With
        
    End If
    
    j = j + 1
Loop

MousePointer = vbDefault

End Sub

Private Sub fill_grid()
Adodc1.RecordSource = "select * from tblenroll order by enrollnumber"
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

lblMessage.Caption = "Working..."
DoEvents
vEnrollNumber = Val(txtEnrollNumber.Text)
vFingerNumber = cmbBackupNumber.Text

vRet = frmMain.CZKEM1.DeleteEnrollData(frmMain.vMachineNumber, vEnrollNumber, frmMain.vMachineNumber, vFingerNumber)
If vRet = True Then
    lblMessage.Caption = "DeleteEnrollData OK"
Else
    frmMain.CZKEM1.GetLastError vErrorCode
    lblMessage.Caption = ErrorPrint(vErrorCode)
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

Private Sub cmd_export_Click()

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Adodc1.Caption = "Record " & Adodc1.Recordset.AbsolutePosition & " of " & Adodc1.Recordset.RecordCount
End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Adodc2.Caption = "Record " & Adodc2.Recordset.AbsolutePosition & " of " & Adodc2.Recordset.RecordCount
End Sub

Private Sub Form_Load()
Me.Width = 11070
Me.Height = 8790

Adodc1.ConnectionString = strConn
Adodc2.ConnectionString = strConn

Call fill_grid
Call fill_grid_device

Call DataGrid1_RowColChange(1, 1)
Call DataGrid2_RowColChange(1, 1)
End Sub
