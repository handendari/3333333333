VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "ATT2000"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Clear Old ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   12
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   240
      Top             =   2880
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   11
      Top             =   3840
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   6240
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   794
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
   Begin VB.OptionButton Option2 
      Caption         =   "Setting"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Get Data"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdFlash 
      Caption         =   "Flash"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer tmrFlash 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Tray"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tray it"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Get Data"
      Height          =   3015
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   7215
      Begin VB.CommandButton Command4 
         Caption         =   "Get Data"
         Height          =   615
         Left            =   2400
         TabIndex        =   10
         Top             =   1320
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Setting"
      Height          =   3015
      Left            =   600
      TabIndex        =   6
      Top             =   720
      Width           =   7215
      Begin VB.CommandButton Command3 
         Caption         =   "&Save"
         Height          =   495
         Left            =   4680
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txt_timer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   720
         Top             =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Timer (Minute)"
         Height          =   195
         Left            =   960
         TabIndex        =   7
         Top             =   1440
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mWndProcNext As Long
Private isSubclassed As Boolean
' Mouse Event messages. I have intentionally only shown how to
' detect one of these messages. It's is easy to implement the others
' into the code though.
' Left button mouse messages
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
' Middle button mouse messages
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209
' Right button mouse messages
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206

'---
'Option Explicit
Private TrayIcon As clsTrayIcon
' Userdefined message. This can be change as you want but remember to
' change it in the class too.
Private Const WM_USER As Long = &H400
Private Const WM_MYHOOK As Long = WM_USER + 1
Dim int_timer As Integer
Dim Cn As New ADODB.Connection


Private Sub Command1_Click()
Set TrayIcon = New clsTrayIcon
   ' Set up the properties for the tray icon and create it.
   With TrayIcon
      Set .Icon = Me.Icon ' Set application icon as SysTray icon
      .ToolTip = "ATT2000"
      .ParentHandle = Me.hwnd
      .trayCreate
      
   End With
Me.Visible = False
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False

cn_text.Execute "drop table a.txt"
cn_text.Execute "create table a.txt(interval text)"
cn_text.Execute "insert into a.txt (interval) values (" & Trim(txt_timer.Text) & ")"

int_interval = Val(txt_timer)
Timer1.Enabled = True
End Sub

Private Sub load_interval()
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

cmd.ActiveConnection = cn_text
rs.CursorLocation = adUseClient
cmd.CommandText = "select interval from a.txt"
rs.Open cmd, , adOpenStatic, adLockBatchOptimistic

If rs.RecordCount > 0 Then
    txt_timer = rs.Fields("interval").Value
    int_interval = Val(txt_timer)
End If
End Sub

Private Sub Command4_Click()
On Error GoTo errCap
Timer1.Enabled = False
    
'Adodc1.RecordSource = "select A.*, B.BADGENUMBER from CHECKINOUT A inner join USERINFO B " _
'& "on A.USERID=B.USERID order by CHECKTIME asc"

Adodc1.RecordSource = "select A.*, B.BADGENUMBER from CHECKINOUT A inner join USERINFO B " _
& "on A.USERID=B.USERID where iif(flag_db=Null,0,flag_db)=0 and " _
& "len(B.badgenumber) = 8 order by CHECKTIME asc"

Adodc1.Refresh

'Adodc1.Recordset.MoveFirst
'MsgBox Adodc1.Recordset.RecordCount

'While Not Adodc1.Recordset.EOF
'    cn_dbf.Execute "insert into USER_INFO" _
'    & "(USERID, BADGENUM, CHECKTIME, CHECKTYPE, VERIFYCODE, SENSORID, ) values (" _
'    & Adodc1.Recordset("USERID").Value & ",'" & Adodc1.Recordset("BADGENUMBER").Value & "'," _
'    & "datetime(" & Format(Adodc1.Recordset("CHECKTIME").Value, "yyyy,mm,dd,hh,nn,ss") & "),'" & Adodc1.Recordset("CHECKTYPE").Value & "'," _
'    & Adodc1.Recordset("VERIFYCODE").Value & ",'" & Adodc1.Recordset("SENSORID").Value & "')"
'    Adodc1.Recordset.MoveNext
'Wend

While Not Adodc1.Recordset.EOF
    cn_dbf.Execute "insert into USR_INFO" _
    & "(USERID, BADGENUM, CHECKDATE, CHECKTYPE, VERIFYCODE, SENSORID, FLAG, CHECKTIME) values (" _
    & Adodc1.Recordset("USERID").Value & ",'" & Adodc1.Recordset("BADGENUMBER").Value & "'," _
    & "datetime(" & Format(Adodc1.Recordset("CHECKTIME").Value, "yyyy,mm,dd,hh,nn,ss") & "),'" & Adodc1.Recordset("CHECKTYPE").Value & "'," _
    & Adodc1.Recordset("VERIFYCODE").Value & ",'" & Adodc1.Recordset("SENSORID").Value _
    & "', '','" & Format(Adodc1.Recordset("CHECKTIME").Value, "hh:nn:ss") & "')"
    
    Cn.Execute "update checkinout set flag_db=1 where userid=" & Adodc1.Recordset.Fields("userid").Value _
    & " and format(checktime,'yyyy-mm-dd hh:nn:ss')='" & Format(Adodc1.Recordset.Fields("checktime").Value, "yyyy-mm-dd hh:nn:ss") & "'"
    
    Adodc1.Recordset.MoveNext
    
    
    
Wend

'Cn.Execute "delete from checkinout where USERID in (select USERID from userinfo)"

'MsgBox "Getting Data finished successfully...", vbInformation, "ATT2000"
Timer1.Enabled = True
Exit Sub
errCap:
MsgBox "Error getting data...", vbCritical, "ATT2000"
Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub SubClass(sHwnd)
   
   Dim lResult As Long
   
   ' Make sure we stop all subclassing
   UnSubClass
   
   mWndProcNext = SetWindowLong(sHwnd, GWL_WNDPROC, AddressOf SubWndProc)
   
   If mWndProcNext Then
      lResult = SetWindowLong(sHwnd, GWL_USERDATA, ObjPtr(Me))
      isSubclassed = True
   End If
   
End Sub

Private Sub UnSubClass()
   ' Stop subclassing and pass all messages to the VB handled WindowProc
   If mWndProcNext Then
      SetWindowLong Me.hwnd, GWL_WNDPROC, mWndProcNext
      isSubclassed = False
   End If
End Sub

Friend Function WindowProc(ByVal sHwnd As Long, ByVal uMsg As Long, _
       ByVal wParam As Long, ByVal lParam As Long) As Long
   ' This is where we will handle all messages. Handle only messages
   ' for the Form (Form1)
   If sHwnd = Me.hwnd Then
      Select Case uMsg
      Case WM_MYHOOK
         ' the lParam of the message holds the message we want
         ' to handle
         Select Case lParam
         Case WM_LBUTTONUP
            ' Left mousebutton was pressed on the icon.
            'MsgBox "Left Button was Pressed"
         Case WM_LBUTTONDBLCLK
            Me.Visible = True
         End Select
      End Select
   End If
   
   ' Pass all messages back to the defaul WindowProc
   WindowProc = CallWindowProc(mWndProcNext, sHwnd, uMsg, _
                wParam, ByVal lParam)
        
End Function

Private Sub Command6_Click()
Cn.BeginTrans
Cn.Execute "DELETE FROM USERINFO WHERE LEN(BADGENUMBER)<8"
Cn.CommitTrans
End Sub

Private Sub Form_Load()
SubClass Me.hwnd
Call Option1_Click
'Call load_interval

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
& App.Path & "\att2000.mdb;Persist Security Info=False"

Cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
& App.Path & "\att2000.mdb;Persist Security Info=False"
Cn.Open

int_timer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call UnSubClass
End Sub

Private Sub Option1_Click()
If Option1 Then
    Frame1.Visible = True
    Frame2.Visible = False
End If
End Sub

Private Sub Option2_Click()
If Option2 Then
    Frame2.Visible = True
    Frame1.Visible = False
End If
End Sub

Private Sub Timer1_Timer()
int_timer = int_timer + 1

If int_timer >= (int_interval * 60) Then
    Call Command4_Click
    int_timer = 0
End If
End Sub

Private Sub Timer2_Timer()
Call Command1_Click
Timer2.Enabled = False
End Sub

