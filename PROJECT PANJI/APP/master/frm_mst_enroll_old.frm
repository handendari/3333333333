VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#5.10#0"; "zkemkeeper.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "Msdatgrd.ocx"
Begin VB.Form frm_mst_enroll_old 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MASTER ENROLL"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10980
   Begin VB.CommandButton Command3 
      Caption         =   "&Set"
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
      Left            =   1560
      Picture         =   "frm_mst_enroll_old.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Unset"
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
      Left            =   2760
      Picture         =   "frm_mst_enroll_old.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   1095
   End
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   375
      Left            =   -120
      OleObjectBlob   =   "frm_mst_enroll_old.frx":0B14
      TabIndex        =   11
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmd_refresh 
      Caption         =   "&Refresh"
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
      Left            =   7080
      Picture         =   "frm_mst_enroll_old.frx":0B38
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmd_import 
      Caption         =   "&Download"
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
      Left            =   360
      Picture         =   "frm_mst_enroll_old.frx":10C2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
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
      Left            =   9600
      Picture         =   "frm_mst_enroll_old.frx":164C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5280
      Top             =   5400
   End
   Begin VB.Frame fra_source 
      Caption         =   "SOURCE (FINGERPRINT DEVICE)"
      Height          =   4695
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4215
      Begin VB.CommandButton cmd_get_data_source 
         Caption         =   "Get Data"
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
         Left            =   2160
         TabIndex        =   6
         Top             =   2040
         Visible         =   0   'False
         Width           =   1515
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
         Left            =   480
         TabIndex        =   5
         Top             =   2040
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtIPAddress_source 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   2
         Text            =   "10.155.9.75"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtPortNo_source 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   1
         Text            =   "4370"
         Top             =   1155
         Width           =   1815
      End
      Begin VB.Label lblIPAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ip Address "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   4
         Top             =   660
         Width           =   1080
      End
      Begin VB.Label lblPortNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Top             =   1155
         Width           =   1200
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5880
      Top             =   4680
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   5880
      TabIndex        =   7
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7646
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
      Caption         =   "MASTER ENROLL"
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
   Begin VB.Line Line1 
      X1              =   360
      X2              =   10560
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line2 
      X1              =   5760
      X2              =   5160
      Y1              =   2520
      Y2              =   1680
   End
   Begin VB.Line Line3 
      X1              =   5160
      X2              =   5760
      Y1              =   3360
      Y2              =   2520
   End
   Begin VB.Line Line4 
      X1              =   4680
      X2              =   5160
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line5 
      X1              =   4680
      X2              =   5160
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line6 
      X1              =   5160
      X2              =   5160
      Y1              =   2880
      Y2              =   3360
   End
   Begin VB.Line Line7 
      X1              =   5160
      X2              =   5160
      Y1              =   1680
      Y2              =   2160
   End
End
Attribute VB_Name = "frm_mst_enroll_old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim glngEnrollData(DATASIZE) As Long
Dim glngEnrollPData As Integer
Dim gbytEnrollData(DATASIZE * 5) As Byte
Dim vEMachineNumber As Integer
Public en_number_buf, jml_en_rec As Long



Public Sub get_data_source()
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

'---- Get Enroll data and save it into database -------------
j = en_number_buf
If j >= lng_start And j <= lng_end Then
    vEnrollNumber = j
        
    vRet = CZKEM1.GetEnrollData(vMachineNumber, vEnrollNumber, vEMachineNumber, _
            vFingerNumber, vPrivilege, glngEnrollData(0), glngEnrollPData)
       
    If vRet Then
        jml_en_rec = jml_en_rec + 1
        If rs.State = 1 Then rs.Close
        rs.Open "select count(*) as jml from m_enroll where enrollnumber = " & vEnrollNumber, CnG, adOpenStatic, adLockReadOnly
        
        Dim edit_rec As Boolean: edit_rec = False
        If rs.Fields("jml").Value > 0 Then
            If set_bookmark_ado(vEnrollNumber) = True Then
                edit_rec = True
            End If
        End If
        
        With Adodc1.Recordset
            If edit_rec = False Then .AddNew
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
End If
End Sub

Private Function set_bookmark_ado(ByVal lng_no As Long) As Boolean
Adodc1.Recordset.MoveFirst

Adodc1.Recordset.Find ("enrollnumber=" & lng_no)   ', 0, adSearchForward, 1)
If Not (Adodc1.Recordset.EOF = True Or Adodc1.Recordset.BOF = True) Then
    set_bookmark_ado = True
Else
    set_bookmark_ado = False
End If
End Function

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_get_data_source_Click()
Call get_data_source
End Sub

Private Sub cmd_import_Click()
Timer1.Enabled = True
End Sub

Public Sub cmd_refresh_Click()
Adodc1.Recordset.Close
Call fill_grid
End Sub

Public Sub cmdConnect_source_Click()
Dim bConn As Boolean

If cmdConnect_source.Caption = "Disconnect" Then
    CZKEM1.EnableDevice vMachineNumber, True
    CZKEM1.disconnect
    cmdConnect_source.Caption = "Connect"

ElseIf cmdConnect_source.Caption = "Connect" Then
    bConn = CZKEM1.Connect_Net(txtIPAddress_source, CLng(txtPortNo_source))
    If bConn Then
        cmdConnect_source.Caption = "Disconnect"
        CZKEM1.EnableDevice vMachineNumber, False
    Else
        'MsgBox "Error connecting to source device...", vbCritical, headerMSG
        Exit Sub
    End If
End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Adodc1.Caption = "Record " & Adodc1.Recordset.AbsolutePosition & " of " & Adodc1.Recordset.RecordCount
End Sub

Private Sub Form_Load()
Me.Height = 6930
Me.Width = 11070

vMachineNumber = 1
vEMachineNumber = 1
Adodc1.ConnectionString = strConn
Call fill_grid

DataGrid1.Columns("NAMA").Visible = False
End Sub

Private Sub fill_grid()
Adodc1.RecordSource = "select * from m_enroll order by enrollnumber asc"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

Call DataGrid1_RowColChange(1, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmdConnect_source.Caption = "Disconnect" Then Call cmdConnect_source_Click
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False

If cmdConnect_source.Caption = "Connect" Then Call cmdConnect_source_Click
If Not cmdConnect_source.Caption = "Disconnect" Then
    MsgBox "Error connecting to source device...", vbCritical, headerMSG
    Exit Sub
End If

frm_progess.Caption = "Downloading..."
en_number_buf = 0
jml_en_rec = 0
frm_progess.Show 1
End Sub

