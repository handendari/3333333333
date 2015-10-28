VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#5.10#0"; "zkemkeeper.dll"
Begin VB.Form frm_trans_log 
   Caption         =   "LOG ATTENDANCE"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14685
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   960
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Height          =   345
      Left            =   1800
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1515
   End
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_trans_log.frx":0000
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   360
      TabIndex        =   1
      Top             =   7440
      Width           =   13935
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
         Left            =   6480
         TabIndex        =   11
         Text            =   "4370"
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   6480
         TabIndex        =   10
         Text            =   "192.168.1.201"
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   3240
         Picture         =   "frm_trans_log.frx":0024
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   4320
         Picture         =   "frm_trans_log.frx":05AE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   1080
         Picture         =   "frm_trans_log.frx":0B38
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   11520
         Picture         =   "frm_trans_log.frx":10C2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5400
         Picture         =   "frm_trans_log.frx":164C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   2160
         Picture         =   "frm_trans_log.frx":1BD6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_download 
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
         Height          =   645
         Left            =   8280
         Picture         =   "frm_trans_log.frx":2160
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Timer timer_grid 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6975
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   12303
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO URUT"
      Columns(0).DataField=   "no_urut"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "TANGGAL"
      Columns(1).DataField=   "tanggal_"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ID FP"
      Columns(2).DataField=   "id_fp"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "KODE KARYAWAN"
      Columns(3).DataField=   "kode_karyawan"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "NAMA KARYAWAN"
      Columns(4).DataField=   "nama_karyawan"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "I/O"
      Columns(5).DataField=   "flag_io_"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "KETERANGAN"
      Columns(6).DataField=   "status_absensi"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1984"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1905"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=3387"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3307"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=3149"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3069"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=3572"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=3493"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=5318"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=5239"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(4)._MinWidth=10"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=1879"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=1799"
      Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(33)=   "Column(5)._MinWidth=54215968"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=4815"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=4736"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(6)._MinWidth=54215968"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "DAFTAR KEHADIRAN"
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.alignment=0,.bgcolor=&H80000002&"
      _StyleDefs(8)   =   ":id=4,.fgcolor=&H80000009&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(9)   =   ":id=4,.strikethrough=0,.charset=0"
      _StyleDefs(10)  =   ":id=4,.fontname=Tahoma"
      _StyleDefs(11)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&H80000002&,.fgcolor=&H80000009&"
      _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&H8000000A&"
      _StyleDefs(24)  =   ":id=14,.fgcolor=&H80000002&"
      _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=98,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=95,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=96,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=97,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=66,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=102,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=99,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=100,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=101,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=110,.parent=13,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=107,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=108,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=109,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=70,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
      _StyleDefs(62)  =   "Named:id=33:Normal"
      _StyleDefs(63)  =   ":id=33,.parent=0"
      _StyleDefs(64)  =   "Named:id=34:Heading"
      _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   ":id=34,.wraptext=-1"
      _StyleDefs(67)  =   "Named:id=35:Footing"
      _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(69)  =   "Named:id=36:Selected"
      _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(71)  =   "Named:id=37:Caption"
      _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(73)  =   "Named:id=38:HighlightRow"
      _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=39:EvenRow"
      _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(77)  =   "Named:id=40:OddRow"
      _StyleDefs(78)  =   ":id=40,.parent=33"
      _StyleDefs(79)  =   "Named:id=41:RecordSelector"
      _StyleDefs(80)  =   ":id=41,.parent=34"
      _StyleDefs(81)  =   "Named:id=42:FilterBar"
      _StyleDefs(82)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frm_trans_log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn_fp As Boolean
Dim cn_dbf As New ADODB.Connection
Dim rs_dbf As New ADODB.Recordset


Private Function clear_all_log() As Boolean
Dim vRet As Boolean
Dim vErrorCode As Long

vRet = CZKEM1.ClearGLog(vMachineNumber)
If vRet Then
    clear_all_log = True
Else
    clear_all_log = False
    'CZKEM1.GetLastError vErrorCode
    'lblMessage.Caption = ErrorPrint(vErrorCode)
End If
End Function

Public Sub download_data_log()
On Error GoTo capErr

Dim vTMachineNumber, vSMachineNumber, vSEnrollNumber As Long
Dim vVerifyMode, vInOutMode As Long
Dim vYear, vMonth, vDay As Long
Dim vHour, vMinute As Long
Dim vErrorCode As Long
Dim vRet As Boolean
Dim i, n As Long

Dim str_buf_dt, str_buf_kode, str_buf_nama As String
Dim str_start_time_ref, str_end_time_ref, str_time As String
Dim int_status_absensi As Integer
Dim rs As New ADODB.Recordset

i = 0
DoEvents
MousePointer = vbHourglass

Do
    vRet = CZKEM1.GetAllGLogData(vMachineNumber, vTMachineNumber, _
                                vSEnrollNumber, vSMachineNumber, vVerifyMode, _
                                vInOutMode, vYear, vMonth, vDay, vHour, vMinute)
    
    If (vRet = False) Then Exit Do
    str_buf_dt = CStr(vYear) & "-" & Format(vMonth, "0#") & "-" & Format(vDay, "0#") & _
            " " & Format(vHour, "0#") & ":" & Format(vMinute, "0#") & ":00"
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from v_ref_log_data where id_fp = " & vSEnrollNumber, CnG, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        str_buf_kode = rs.Fields("kode_karyawan").Value
        str_buf_nama = rs.Fields("nama_karyawan").Value
        
        ' === parsing time, check & sign it to "status_absensi" as int 0,1 (late or not)
        str_time = Format(vHour, "0#") & ":" & Format(vMinute, "0#")
        str_start_time_ref = rs.Fields("start_time").Value
        str_end_time_ref = rs.Fields("end_time").Value
        
        If vInOutMode = 0 Then
            If str_time <= str_start_time_ref Then
                int_status_absensi = 0
            Else
                int_status_absensi = 1
            End If
            
        Else
            If str_time >= str_end_time_ref Then
                int_status_absensi = 0
            Else
                int_status_absensi = 1
            End If
        End If
    
    Else
        str_buf_kode = ""
        str_buf_nama = ""
        'MsgBox "Tidak ditemukan data karyawan dengan ID FP = " & vSEnrollNumber, vbCritical, "No Data Found"
        'Exit Sub
    End If
        
    
    CnG.Execute "insert into h_att_karyawan " _
    & "(tanggal, kode_karyawan, nama_karyawan, id_fp, flag_io, status_absensi) values ('" _
    & str_buf_dt & "','" & str_buf_kode & "','" & str_buf_nama & "'," & vSEnrollNumber _
    & "," & vInOutMode & "," & int_status_absensi & ")"
    
    
    With rs_dbf
        .AddNew
        
        .Fields("userid").Value = 1
        .Fields("badgenum").Value = vSEnrollNumber
        .Fields("checkdate").Value = str_buf_dt
        .Fields("checktype").Value = IIf(vInOutMode = 0, "I", "O")
        .Fields("verifycode").Value = 0
        .Fields("sensorid").Value = 1
        .Fields("checktime").Value = Right(str_buf_dt, 8)
        
        .Update
    End With
    
    i = i + 1
Loop

MousePointer = vbDefault

MsgBox "Jumlah data ter-download = " & i, vbInformation, headerMSG
Exit Sub

capErr:
MsgBox err.Number
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

Private Function connect() As Boolean
If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, True
    CZKEM1.disconnect
End If

cn_fp = CZKEM1.Connect_Net(txtIPAddress_source, CLng(txtPortNo_source))

If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, False
    connect = True
Else
    connect = False
    Exit Function
End If
End Function

Private Function disconnect() As Boolean
If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, True
    CZKEM1.disconnect
End If
End Function

Public Sub cmd_download_Click()
If connect = False Then
    MsgBox "Error connecting to source device...", vbCritical, headerMSG
    Exit Sub
End If

If rs_dbf.State = 1 Then rs_dbf.Close
rs_dbf.Open "select * from usr_info where userid=-777", cn_dbf, adOpenKeyset, adLockOptimistic

Call download_data_log
If clear_all_log = False Then
    MsgBox "Error Deleting log from device...", vbCritical, headerMSG
End If
Call disconnect

Call load_data_grid
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = strConn

cn_dbf.ConnectionString = "Provider=VFPOLEDB.1;Data Source=" _
& App.Path & ";Password=;Collating Sequence=MACHINE"
cn_dbf.Open

cn_fp = False
vMachineNumber = 1
vEMachineNumber = 1
Call load_data_grid
End Sub

Private Sub load_data_grid()
timer_grid.Enabled = True
End Sub

Private Sub timer_grid_Timer()
Adodc1.RecordSource = "select *, format(tanggal,'dd-MM-yyyy hh:nn:ss') as tanggal_, iif(flag_io=0,'I','O') as flag_io_ " _
& "from h_att_karyawan order by tanggal desc, kode_karyawan asc"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1

timer_grid.Enabled = False
End Sub
