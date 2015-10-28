VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_import_attendance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IMPORT ATTENDANCE DATA"
   ClientHeight    =   7815
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   14700
   Icon            =   "frm_trans_import_att.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   4650
      TabIndex        =   5
      Top             =   1440
      Width           =   5655
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   270
         TabIndex        =   6
         Top             =   2100
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   300
         TabIndex        =   8
         Top             =   1740
         Width           =   2475
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Processing, Please Wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   300
         Width           =   2730
      End
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      Picture         =   "frm_trans_import_att.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmd_search 
      Caption         =   "&Browse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "frm_trans_import_att.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      Picture         =   "frm_trans_import_att.frx":0B20
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmd_refresh 
      Cancel          =   -1  'True
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Picture         =   "frm_trans_import_att.frx":10AA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7200
      Top             =   6960
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6495
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   11456
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "DATE"
      Columns(0).DataField=   "att_date"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "IP ADDRESS"
      Columns(1).DataField=   "ip_address"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ENROLL NO."
      Columns(2).DataField=   "enrollnumber"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "EMP. CODE"
      Columns(3).DataField=   "employee_code"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4022"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3942"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3810"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3731"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3466"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3387"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8708"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(2)._MinWidth=80000960"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=4736"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=4657"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=8708"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(3)._MinWidth=79999936"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "LOG ATTENDANCE"
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
      _StyleDefs(21)  =   "Splits(0).Style:id=123,.parent=1"
      _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=132,.parent=4,.bgcolor=&H80000002&"
      _StyleDefs(23)  =   ":id=132,.fgcolor=&H80000009&"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=124,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(25)  =   ":id=124,.fgcolor=&H80000002&"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=125,.parent=3"
      _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=126,.parent=5"
      _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=128,.parent=6"
      _StyleDefs(29)  =   "Splits(0).EditorStyle:id=127,.parent=7"
      _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=129,.parent=8"
      _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=130,.parent=9"
      _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=131,.parent=10"
      _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=133,.parent=11"
      _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=134,.parent=12"
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=146,.parent=123"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=143,.parent=124"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=144,.parent=125"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=145,.parent=127"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=150,.parent=123"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=147,.parent=124"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=148,.parent=125"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=149,.parent=127"
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=162,.parent=123,.locked=-1"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=159,.parent=124"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=160,.parent=125"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=161,.parent=127"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=166,.parent=123,.locked=-1"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=163,.parent=124"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=164,.parent=125"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=165,.parent=127"
      _StyleDefs(51)  =   "Named:id=33:Normal"
      _StyleDefs(52)  =   ":id=33,.parent=0"
      _StyleDefs(53)  =   "Named:id=34:Heading"
      _StyleDefs(54)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   ":id=34,.wraptext=-1"
      _StyleDefs(56)  =   "Named:id=35:Footing"
      _StyleDefs(57)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(58)  =   "Named:id=36:Selected"
      _StyleDefs(59)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(60)  =   "Named:id=37:Caption"
      _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(62)  =   "Named:id=38:HighlightRow"
      _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(64)  =   "Named:id=39:EvenRow"
      _StyleDefs(65)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(66)  =   "Named:id=40:OddRow"
      _StyleDefs(67)  =   ":id=40,.parent=33"
      _StyleDefs(68)  =   "Named:id=41:RecordSelector"
      _StyleDefs(69)  =   ":id=41,.parent=34"
      _StyleDefs(70)  =   "Named:id=42:FilterBar"
      _StyleDefs(71)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frm_trans_import_attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBound_m As New ADODB.Recordset
Dim rsBound_d As New ADODB.Recordset
Dim rsBound_o As New ADODB.Recordset
Dim rsBound_b As New ADODB.Recordset

Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns


Private Sub fill_grid_excel_m(str_file_name As String)
Dim strWorksheet, strWorksheet_m, strWorksheet_d As String
'Screen.MousePointer = vbHourglass
'DoEvents
strWorksheet = "h_log_attendance"
strWorksheet_m = "tm_bom": strWorksheet_d = "td_bom"

Adodc1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
& str_file_name & ";Extended Properties=Excel 8.0"

Adodc1.RecordSource = "select * from [" & strWorksheet & "$] order by att_date asc"
Adodc1.Refresh
TDBGrid1.DataSource = Adodc1
End Sub

Private Sub fill_grid_excel_d(str_file_name As String)
If Trim(str_file_name) = "" Then Exit Sub
Dim strWorksheet_m, strWorksheet_d As String

strWorksheet_m = "tm_bom": strWorksheet_d = "td_bom"

Adodc2.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
& str_file_name & ";Extended Properties=Excel 8.0"

Adodc2.RecordSource = "select * from [" & strWorksheet_d & "$] where no_bom = '" _
& TDBGrid1.Columns("no_bom").Value & "' and no_urut_revisi = " _
& TDBGrid1.Columns("no_urut_revisi").Value & " order by kode_jenis_bom, nama_barang"
Adodc2.Refresh
TDBGrid2.DataSource = Adodc2
End Sub

Private Sub fill_grid_excel_o(str_file_name As String)
If Trim(str_file_name) = "" Then Exit Sub
Dim strWorksheet_o As String

strWorksheet_o = "td_outsourcing"

Adodc3.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
& str_file_name & ";Extended Properties=Excel 8.0"

Adodc3.RecordSource = "select * from [" & strWorksheet_o & "$] where no_bom = '" _
& TDBGrid1.Columns("no_bom").Value & "' and no_urut_revisi = " _
& TDBGrid1.Columns("no_urut_revisi").Value & " order by kode_jenis_outsource, nama_outsource"
Adodc3.Refresh
TDBGrid3.DataSource = Adodc3
End Sub

Private Sub fill_grid_excel_b(str_file_name As String)
If Trim(str_file_name) = "" Then Exit Sub
Dim strWorksheet_b As String

strWorksheet_b = "td_biaya_umum"

Adodc4.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
& str_file_name & ";Extended Properties=Excel 8.0"

Adodc4.RecordSource = "select * from [" & strWorksheet_b & "$] where no_bom = '" _
& TDBGrid1.Columns("no_bom").Value & "' and no_urut_revisi = " _
& TDBGrid1.Columns("no_urut_revisi").Value & " order by kode_biaya_umum"
Adodc4.Refresh
TDBGrid4.DataSource = Adodc4
End Sub

Private Sub cmd_refresh_Click()
If CommonDialog1.FileName <> "" Then
    Call fill_grid_excel_m(CommonDialog1.FileName)
End If
End Sub

Private Sub cmd_search_Click()
CommonDialog1.Filter = "XLS|*.xls"
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowOpen

If CommonDialog1.FileName <> "" Then
    Call fill_grid_excel_m(CommonDialog1.FileName)
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Frame1.Visible = False
'If rsBound_m.State = 1 Then rsBound_m.Close
'If rsBound_d.State = 1 Then rsBound_d.Close
'If rsBound_o.State = 1 Then rsBound_o.Close
'If rsBound_b.State = 1 Then rsBound_b.Close
'
'Call ExecQuery(rsBound_m, "select * from tm_bom where no_bom='άφ'", 1)
'Call ExecQuery(rsBound_d, "select * from td_bom where no_bom='άφ'", 1)
'Call ExecQuery(rsBound_o, "select * from td_outsourcing where no_bom='άφ'", 1)
'Call ExecQuery(rsBound_b, "select * from td_biaya_umum where no_bom='άφ'", 1)
'
'SSTab1.Tab = 0
End Sub

Private Sub TDBGrid1_FilterChange()
On Error GoTo ErrHandler

Dim i As Integer

Set Cols = TDBGrid1.Columns
i = TDBGrid1.Col
TDBGrid1.HoldFields

Adodc1.Recordset.Filter = getFilter()
TDBGrid1.Col = i
TDBGrid1.EditActive = True

TDBGrid1.SelStart = _
    Len(TDBGrid1.Columns(i).FilterText)
If TDBGrid1.ApproxCount < 1 Then
    Call clear_filter
    TDBGrid1.Col = i
End If
Exit Sub

ErrHandler:
'MsgBox Err.Source & ":" & vbCrLf & Err.Description
MsgBox "Tidak ada data pada kolom ini " & vbCr _
& "atau data filter yang anda inputkan tidak valid", vbCritical, "Filtering not allowed"
Call clear_filter
End Sub

Private Function getFilter() As String
Dim tmp As String
Dim n As Integer

For Each Col In Cols
    If Trim(Col.FilterText) <> "" Then
        n = n + 1
        If n > 1 Then
            tmp = tmp & " AND "
        End If
        
        'If TDBGrid1.Columns(TDBGrid1.Col).Caption = "SATUAN" Then
        '    tmp = tmp & Col.DataField = " & Col.FilterText "
        'Else
            tmp = tmp & Col.DataField & " LIKE '" & Col.FilterText & "*'"
        'End If
    End If
Next Col
getFilter = tmp
End Function

Private Sub clear_filter()
For Each Col In TDBGrid1.Columns
    Col.FilterText = ""
Next Col
Adodc1.Recordset.Filter = adFilterNone
End Sub

Private Sub TDBGrid1_FormatText _
(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBGrid1.Columns(ColIndex).Caption = "TANGGAL" Then
    Value = Format(Value, "dd-mm-yyyy")
End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If TDBGrid1.Row + 1 = LastRow Then Exit Sub
'
'Call fill_grid_excel_d(CommonDialog1.FileName)
'Call fill_grid_excel_o(CommonDialog1.FileName)
'Call fill_grid_excel_b(CommonDialog1.FileName)
End Sub

Private Function insert_new_data() As Boolean
On Error GoTo err_handle

Dim rs1 As New ADODB.Recordset

If rs1.State = 1 Then rs1.Close
rs1.Open "select * from m_employee where employee_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic

CnG.BeginTrans

With rs1
    .AddNew
    
    
    .Fields("employee_code").Value = Adodc1.Recordset.Fields("employee_code").Value
    '--------------------------------
    .Fields("employee_name").Value = Adodc1.Recordset.Fields("employee_name").Value
    .Fields("employee_nick_name").Value = Adodc1.Recordset.Fields("employee_nick_name").Value
    .Fields("division_code").Value = Adodc1.Recordset.Fields("division_code").Value
    .Fields("division_name").Value = Adodc1.Recordset.Fields("division_name").Value
    .Fields("department_code").Value = Adodc1.Recordset.Fields("department_code").Value
    .Fields("department_name").Value = Adodc1.Recordset.Fields("department_name").Value
    .Fields("company_code").Value = Adodc1.Recordset.Fields("company_code").Value
    .Fields("company_name").Value = Adodc1.Recordset.Fields("company_name").Value
    '.Fields("date_of_birth").Value = Adodc1.Recordset.Fields("date_of_birth").Value
    .Fields("date_of_birth").Value = IIf(IsDate(Adodc1.Recordset.Fields("date_of_birth").Value) = True, Adodc1.Recordset.Fields("date_of_birth").Value, vbNull)
    
    .Fields("place_of_birth").Value = Adodc1.Recordset.Fields("place_of_birth").Value
    .Fields("sex").Value = Adodc1.Recordset.Fields("sex").Value
    .Fields("religion").Value = Adodc1.Recordset.Fields("religion").Value
    .Fields("marital_status").Value = Adodc1.Recordset.Fields("marital_status").Value
    .Fields("number_of_children").Value = Adodc1.Recordset.Fields("number_of_children").Value
    .Fields("address").Value = Adodc1.Recordset.Fields("address").Value
    .Fields("email").Value = Adodc1.Recordset.Fields("email").Value
    .Fields("npwp").Value = Adodc1.Recordset.Fields("npwp").Value
    .Fields("phone_number").Value = Adodc1.Recordset.Fields("phone_number").Value
    
    .Fields("last_education_code").Value = Adodc1.Recordset.Fields("last_education_code").Value
    .Fields("last_education_code_other").Value = Adodc1.Recordset.Fields("last_education_code_other").Value
    .Fields("last_education_name").Value = Adodc1.Recordset.Fields("last_education_name").Value
    .Fields("last_education_pass").Value = IIf(IsDate(Adodc1.Recordset.Fields("last_education_pass").Value) = True, Adodc1.Recordset.Fields("last_education_pass").Value, vbNull)
    
    .Fields("last_employment_name").Value = Adodc1.Recordset.Fields("last_employment_name").Value
    .Fields("last_employment_date").Value = IIf(IsDate(Adodc1.Recordset.Fields("last_employment_date").Value) = True, Adodc1.Recordset.Fields("last_employment_date").Value, vbNull)
    
    .Fields("last_employment_title").Value = Adodc1.Recordset.Fields("last_employment_title").Value
    
    .Fields("start_working").Value = IIf(IsDate(Adodc1.Recordset.Fields("start_working").Value) = True, Adodc1.Recordset.Fields("start_working").Value, vbNull)
    
    .Fields("date_of_appointment").Value = IIf(IsDate(Adodc1.Recordset.Fields("date_of_appointment").Value) = True, Adodc1.Recordset.Fields("date_of_appointment").Value, vbNull)
    
    .Fields("title_code").Value = Adodc1.Recordset.Fields("title_code").Value
    .Fields("title_name").Value = Adodc1.Recordset.Fields("title_name").Value
    
    .Fields("level1").Value = Adodc1.Recordset.Fields("level1").Value
    .Fields("level2").Value = Adodc1.Recordset.Fields("level2").Value
    .Fields("bank_name").Value = Adodc1.Recordset.Fields("bank_name").Value
    .Fields("bank_account").Value = Adodc1.Recordset.Fields("bank_account").Value
    
    .Fields("flag_shiftable").Value = Adodc1.Recordset.Fields("flag_shiftable").Value
    .Fields("flag_active").Value = Adodc1.Recordset.Fields("flag_active").Value
    .Fields("description").Value = Adodc1.Recordset.Fields("description").Value
    .Fields("end_working").Value = IIf(IsDate(Adodc1.Recordset.Fields("end_working").Value) = True, Adodc1.Recordset.Fields("end_working").Value, vbNull)
    .Fields("reason").Value = Adodc1.Recordset.Fields("reason").Value
    
    .Update
End With

CnG.CommitTrans
insert_new_data = True

Exit Function

err_handle:
insert_new_data = False
CnG.RollbackTrans
End Function

Private Sub CmdSave_Click()
Dim rsabsen As New ADODB.Recordset
Dim strSQL As String
Dim clsAtt As New clsInsertAttendance
Dim i, j, dep, div, comp, Ttl, log, log_false, link, link_false As Integer

j = 0
dep = 0
div = 0
comp = 0
Ttl = 0
log = 0
log_false = 0
link = 0
link_false = 0

If Not TDBGrid1.ApproxCount > 0 Or Not TDBGrid1.Bookmark > 0 Then
    MsgBox "No data selected!", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Are you sure to import all data?", vbOKCancel, headerMSG)
If Not i = vbOK Then Exit Sub

ProgressBar1.Value = 0

If Adodc1.Recordset.RecordCount > 0 Then Adodc1.Recordset.MoveFirst

For i = 1 To Adodc1.Recordset.RecordCount
    'MsgBox Adodc1.Recordset.Fields(0).Value
    
'    If insert_new_data_comp Then
'        comp = comp + 1
'    End If
'    If insert_new_data_dept Then
'        dep = dep + 1
'    End If
'    If insert_new_data_div Then
'        div = div + 1
'    End If
'    If insert_new_data_title Then
'        Ttl = Ttl + 1
'    End If
'
'    If insert_new_data Then
'        j = j + 1
'    End If

    Frame1.Visible = True
    ProgressBar1.Max = Adodc1.Recordset.RecordCount
    ProgressBar1.Value = ProgressBar1.Value + 1
    
    Label5.Caption = "Verifying Data, Please Wait......"
    Label1.Caption = "Data Ke " & i & " dari " & Adodc1.Recordset.RecordCount
    
'    Call clsAtt.Insert_h_attendance(Format(Adodc1.Recordset.Fields("att_date").Value, "yyyy-MM-dd hh:mm:ss"), _
'        Adodc1.Recordset.Fields("ip_address").Value, Adodc1.Recordset.Fields("enrollnumber").Value)
    
    If insert_new_link Then
        link = link + 1
    Else
        link_false = link_false + 1
    End If
    
    If insert_new_log Then
        log = log + 1
    Else
        log_false = log_false + 1
    End If
        
    Adodc1.Recordset.MoveNext
    DoEvents
Next i

    
    strSQL = "SELECT * FROM ( " & _
        "SELECT DATE(a.att_date) tgl_att,a.employee_code,MIN(att_date) check_in,MAX(att_date) check_out,a.enrollnumber,a.ip_address," & _
            "((HOUR( TIMEDIFF(MAX(att_date),MIN(att_date)))*60) + MINUTE(TIMEDIFF(MAX(att_date),MIN(att_date)) )) selisih " & _
        "FROM h_log_attendance a LEFT JOIN m_employee b ON a.employee_code = b.employee_code " & _
        "WHERE a.sdhupdate = 0 GROUP BY 1,2)xx " & _
        "WHERE selisih > 30"
    
    rsabsen.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rsabsen.RecordCount > 0 Then
        Dim insertAtt As New clsInsertAtt_New
        ProgressBar1.Value = 0
        
        rsabsen.MoveFirst
        While Not rsabsen.EOF
            ProgressBar1.Max = rsabsen.RecordCount
            ProgressBar1.Value = ProgressBar1.Value + 1
            
            Label5.Caption = "Saving Data, Please Wait......"
            Label1.Caption = "Data Ke " & ProgressBar1.Value & " dari " & rsabsen.RecordCount
        
            Call insertAtt.Insert_h_attendance(Format(rsabsen!tgl_att, "yyyy-MM-dd"), rsabsen!employee_code, _
                Format(rsabsen!check_in, "yyyy-MM-dd hh:mm:ss"), Format(rsabsen!check_out, "yyyy-MM-dd hh:mm:ss"), _
                rsabsen!ip_address, rsabsen!enrollnumber)
            rsabsen.MoveNext
        Wend
    End If
    
    '++++++++++++++++
    strSQL = "UPDATE h_log_attendance SET sdhupdate = 1 WHERE sdhupdate = 0"
    CnG.Execute strSQL
    '++++++++++++++++
    
    Frame1.Visible = False
    
    MsgBox log & " log data are successfully import!" & vbCrLf _
        & link & " link data are successfully import!" & vbCrLf _
        & log_false & " log are unsuccessfully import!", vbInformation, headerMSG

    Call frm_mst_employee.load_data_company
    Call frm_mst_employee.load_data_employee
Unload Me
End Sub

Private Function insert_new_link() As Boolean
On Error GoTo err_handle
Dim strSQL As String

'Dim rs1 As New ADODB.Recordset
'
'If rs1.State = 1 Then rs1.Close
'rs1.Open "select * from m_enroll_link where enrollnumber=-77", CnG, adOpenKeyset, adLockOptimistic


'With rs1
'    .AddNew
'
'    .Fields("ip_address").Value = Adodc1.Recordset.Fields("ip_address").Value
'    .Fields("enrollnumber").Value = Adodc1.Recordset.Fields("enrollnumber").Value
'    .Fields("employee_code").Value = Adodc1.Recordset.Fields("employee_code").Value
''    .Fields("city_name").Value = Adodc1.Recordset.Fields("city_name").Value
''    .Fields("phone_number").Value = Adodc1.Recordset.Fields("phone_number").Value
''    .Fields("fax_number").Value = Adodc1.Recordset.Fields("fax_number").Value
''    .Fields("web_address").Value = Adodc1.Recordset.Fields("web_address").Value
''    .Fields("email_address").Value = Adodc1.Recordset.Fields("email_address").Value
''    .Fields("npwp").Value = Adodc1.Recordset.Fields("npwp").Value
''    .Fields("entry_date").Value = Now 'Adodc1.Recordset.Fields("entry_date").Value
'
'    .Update
'End With

    CnG.BeginTrans
        strSQL = "INSERT INTO m_enroll_link (ip_address,enrollnumber,employee_code) " & _
            "VALUES " & _
            "('" & Adodc1.Recordset.Fields("ip_address").Value & "'," & _
            "'" & Adodc1.Recordset.Fields("enrollnumber").Value & "'," & _
            "'" & Adodc1.Recordset.Fields("employee_code").Value & "')"
        CnG.Execute strSQL
        
    CnG.CommitTrans
    insert_new_link = True
    
    Exit Function

err_handle:
   ' MsgBox "Error : " & err.Description
    insert_new_link = False
    CnG.RollbackTrans
End Function


Private Function insert_new_log() As Boolean
On Error GoTo err_handle
Dim strSQL As String

'Dim rs1 As New ADODB.Recordset

'If rs1.State = 1 Then rs1.Close
'rs1.Open "select * from h_log_attendance where att_number=-77", CnG, adOpenKeyset, adLockOptimistic
'
'CnG.BeginTrans
'
'With rs1
'    .AddNew
'
'    .Fields("att_date").Value = Format(Adodc1.Recordset.Fields("att_date").Value, "yyyy-mm-dd")
'    .Fields("ip_address").Value = Adodc1.Recordset.Fields("ip_address").Value
'    .Fields("enrollnumber").Value = Adodc1.Recordset.Fields("enrollnumber").Value
'    .Fields("employee_code").Value = Adodc1.Recordset.Fields("employee_code").Value
''    .Fields("postal_code").Value = Adodc1.Recordset.Fields("postal_code").Value
''    .Fields("city_name").Value = Adodc1.Recordset.Fields("city_name").Value
''    .Fields("phone_number").Value = Adodc1.Recordset.Fields("phone_number").Value
''    .Fields("fax_number").Value = Adodc1.Recordset.Fields("fax_number").Value
''    .Fields("web_address").Value = Adodc1.Recordset.Fields("web_address").Value
''    .Fields("email_address").Value = Adodc1.Recordset.Fields("email_address").Value
''    .Fields("npwp").Value = Adodc1.Recordset.Fields("npwp").Value
'    .Fields("entry_date").Value = Now 'Adodc1.Recordset.Fields("entry_date").Value
'
'    .Update
'End With

    CnG.BeginTrans
        strSQL = "INSERT INTO h_log_attendance (att_date,ip_address,enrollnumber,employee_code,sdhupdate) " & _
            "VALUES " & _
            "('" & Format(Adodc1.Recordset.Fields("att_date").Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
            "'" & Adodc1.Recordset.Fields("ip_address").Value & "'," & _
            "'" & Adodc1.Recordset.Fields("enrollnumber").Value & "'," & _
            "'" & Adodc1.Recordset.Fields("employee_code").Value & "',0)"
        CnG.Execute strSQL
    CnG.CommitTrans
    insert_new_log = True

Exit Function

err_handle:
    'MsgBox "Error : " & err.Description
    insert_new_log = False
    CnG.RollbackTrans
End Function

Private Function insert_new_data_comp() As Boolean
On Error GoTo err_handle

Dim rs1 As New ADODB.Recordset

If rs1.State = 1 Then rs1.Close
rs1.Open "select * from m_company where company_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic

CnG.BeginTrans

With rs1
    .AddNew
    
    
    .Fields("company_code").Value = Adodc1.Recordset.Fields("company_code").Value
    .Fields("company_name").Value = Adodc1.Recordset.Fields("company_name").Value
'    .Fields("address").Value = Adodc1.Recordset.Fields("address").Value
'    .Fields("postal_code").Value = Adodc1.Recordset.Fields("postal_code").Value
'    .Fields("city_name").Value = Adodc1.Recordset.Fields("city_name").Value
'    .Fields("phone_number").Value = Adodc1.Recordset.Fields("phone_number").Value
'    .Fields("fax_number").Value = Adodc1.Recordset.Fields("fax_number").Value
'    .Fields("web_address").Value = Adodc1.Recordset.Fields("web_address").Value
'    .Fields("email_address").Value = Adodc1.Recordset.Fields("email_address").Value
'    .Fields("npwp").Value = Adodc1.Recordset.Fields("npwp").Value
    
    .Update
End With

CnG.CommitTrans
insert_new_data_comp = True

Exit Function

err_handle:
insert_new_data_comp = False
CnG.RollbackTrans
End Function

Private Function insert_new_data_dept() As Boolean
On Error GoTo err_handle

Dim rs1 As New ADODB.Recordset

If rs1.State = 1 Then rs1.Close
rs1.Open "select * from m_department where department_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic

CnG.BeginTrans

With rs1
    .AddNew
    
    
    .Fields("company_code").Value = Adodc1.Recordset.Fields("company_code").Value
    .Fields("department_code").Value = Adodc1.Recordset.Fields("department_code").Value
    .Fields("department_name").Value = Adodc1.Recordset.Fields("department_name").Value
'    .Fields("description").Value = Adodc1.Recordset.Fields("description").Value
    
    .Update
End With

CnG.CommitTrans
insert_new_data_dept = True

Exit Function

err_handle:
insert_new_data_dept = False
CnG.RollbackTrans
End Function

Private Function insert_new_data_div() As Boolean
On Error GoTo err_handle

Dim rs1 As New ADODB.Recordset

If rs1.State = 1 Then rs1.Close
rs1.Open "select * from m_division where division_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic

CnG.BeginTrans

With rs1
    .AddNew
    
    
    .Fields("company_code").Value = Adodc1.Recordset.Fields("company_code").Value
    .Fields("department_code").Value = Adodc1.Recordset.Fields("department_code").Value
    .Fields("department_name").Value = Adodc1.Recordset.Fields("department_name").Value
    .Fields("division_code").Value = Adodc1.Recordset.Fields("division_code").Value
    .Fields("division_name").Value = Adodc1.Recordset.Fields("division_name").Value
'    .Fields("description").Value = Adodc1.Recordset.Fields("description").Value
    
    .Update
End With

CnG.CommitTrans
insert_new_data_div = True

Exit Function

err_handle:
insert_new_data_div = False
CnG.RollbackTrans
End Function

Private Function insert_new_data_title() As Boolean
On Error GoTo err_handle

Dim rs1 As New ADODB.Recordset

If rs1.State = 1 Then rs1.Close
rs1.Open "select * from m_title where title_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic

CnG.BeginTrans

With rs1
    .AddNew
    
    
    .Fields("title_code").Value = Adodc1.Recordset.Fields("title_code").Value
    .Fields("title_name").Value = Adodc1.Recordset.Fields("title_name").Value
'    .Fields("default_shiftable").Value = Adodc1.Recordset.Fields("default_shiftable").Value
'    .Fields("description").Value = Adodc1.Recordset.Fields("description").Value
    
    .Update
End With

CnG.CommitTrans
insert_new_data_title = True

Exit Function

err_handle:
insert_new_data_title = False
CnG.RollbackTrans
End Function

