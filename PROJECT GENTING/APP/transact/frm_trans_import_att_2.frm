VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_trans_import_attendance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IMPORT ATTENDANCE DATA"
   ClientHeight    =   8205
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   14700
   Icon            =   "frm_trans_import_att_2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7500
      Top             =   7620
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
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
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   4650
      TabIndex        =   4
      Top             =   1860
      Width           =   5655
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   270
         TabIndex        =   5
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
         TabIndex        =   7
         Top             =   1740
         Visible         =   0   'False
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
         TabIndex        =   6
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
      Picture         =   "frm_trans_import_att_2.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7260
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
      Picture         =   "frm_trans_import_att_2.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7260
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
      Picture         =   "frm_trans_import_att_2.frx":0B20
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7260
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
      Picture         =   "frm_trans_import_att_2.frx":10AA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7260
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prj_genting.LynxGrid LynxGrid1 
      Height          =   6435
      Left            =   270
      TabIndex        =   8
      Top             =   690
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   11351
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorSel    =   12937777
      ForeColorSel    =   16777215
      CustomColorFrom =   16572875
      CustomColorTo   =   14722429
      GridColor       =   16367254
      FocusRectColor  =   9895934
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   -1  'True
      AllowColumnResizing=   -1  'True
      ColumnSort      =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "IMPORT KEHADIRAN"
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
      Left            =   5895
      TabIndex        =   9
      Top             =   0
      Width           =   3285
   End
End
Attribute VB_Name = "frm_trans_import_attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fill_grid_excel_m(str_file_name As String)
Dim strWorksheet As String
Dim total_OT As Double

    Me.MousePointer = vbHourglass
    'DoEvents
    strWorksheet = "att"
    'strWorksheet_m = "tm_bom": strWorksheet_d = "td_bom"
    
    Adodc1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & str_file_name & ";Extended Properties=Excel 8.0"
    
    Adodc1.RecordSource = "select * from [" & strWorksheet & "$]"
    Adodc1.Refresh
    
    LynxGrid1.Clear
    With Adodc1.Recordset
        If .RecordCount > 0 Then
            .MoveFirst
            While Not .EOF
                If IsNull(.Fields(0)) = False Then
                    total_OT = (1.5 * IIf(IsNull(.Fields(5)), 0, .Fields(5))) + (2 * IIf(IsNull(.Fields(6)), 0, .Fields(6))) + (3 * IIf(IsNull(.Fields(7)), 0, .Fields(7))) + (4 * IIf(IsNull(.Fields(8)), 0, .Fields(8)))
                    
                    LynxGrid1.AddItem .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2) & vbTab & .Fields(3) _
                    & vbTab & .Fields(4) & vbTab & .Fields(5) & vbTab & .Fields(6) & vbTab & .Fields(7) & vbTab & .Fields(8) & vbTab & total_OT
                End If
                .MoveNext
            Wend
            LynxGrid1.Redraw = True
        End If
    End With
    Me.MousePointer = vbNormal
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
    createGrid
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "Tanggal", 1200, lgAlignCenterCenter, lgDate, "yyyy-MM-dd", , , , , True
      .AddColumn "NIK", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Employee Name", 3500, , , , , , , , True
      .AddColumn "Status", 1200, lgAlignCenterCenter, , , , , , , , True
      .AddColumn "HK", 1200, lgAlignRightCenter, , , , , , , True
      .AddColumn "X 1.5", 1200, lgAlignRightCenter, lgNumeric, "#,##.##", , , , , , True
      .AddColumn "X 2.0", 1200, lgAlignRightCenter, lgNumeric, "#,##.##", , , , , True
      .AddColumn "X 3.0", 1200, lgAlignRightCenter, lgNumeric, "#,##.##", , , , , , True
      .AddColumn "X 4.0", 1200, lgAlignRightCenter, lgNumeric, "#,##.##", , , , , True
      .AddColumn "TOTAL OT", 1200, lgAlignRightCenter, lgNumeric, "#,##.##", , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
End Sub

Private Sub CmdSave_Click()
Dim strsql As String
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

i = MsgBox("Are you sure to import all data?", vbOKCancel, headerMSG)
If Not i = vbOK Then Exit Sub

ProgressBar1.Value = 0

    If LynxGrid1.Rows > 0 Then
        Dim insertAtt As New clsInsertAtt_New_2
        Dim tgl_att As String, kdkaryawan As String, enroll_number As String
        
        ProgressBar1.Value = 0
        With LynxGrid1
            For i = 0 To .Rows - 1
                ProgressBar1.Max = .Rows
                ProgressBar1.Value = ProgressBar1.Value + 1
                
                
                Label5.Caption = "Saving Data, Please Wait......"
                Label1.Caption = "Data Ke " & ProgressBar1.Value & " dari " & i
                
                
                Call insertAtt.Insert_h_attendance(Format(.CellText(i, 0), "yyyy-MM-dd"), .CellText(i, 1), "1", .CellText(i, 3), _
                    .CellValue(i, 4), .CellValue(i, 5), .CellValue(i, 6), .CellValue(i, 7), .CellValue(i, 8), .CellValue(i, 9), _
                    "0", .CellText(i, 1), "")
      
'      .AddColumn "Tanggal", 1200, lgAlignCenterCenter, lgDate, "yyyy-MM-dd", , , , , True
'      .AddColumn "NIK", 1500, lgAlignCenterCenter, , , , , , , True
'      .AddColumn "Employee Name", 3500, , , , , , , , True
'      .AddColumn "HK", 1200, lgAlignRightCenter, lgNumeric, "#,##.##", , , , , True
'      .AddColumn "X 1.5", 1200, lgAlignRightCenter, lgNumeric, "#,##.##", , , , , , True
'      .AddColumn "X 2.0", 1200, lgAlignRightCenter, lgNumeric, "#,##.##", , , , , True
'      .AddColumn "X 3.0", 1200, lgAlignRightCenter, lgNumeric, "#,##.##", , , , , , True
'      .AddColumn "X 4.0", 1200, lgAlignRightCenter, lgNumeric, "#,##.##", , , , , True
'      .AddColumn "TOTAL OT", 1200, lgAlignRightCenter, lgNumeric, "#,##.##", , , , , True
            Next
        End With
    End If
    
    Frame1.Visible = False
    
'    MsgBox log & " log data are successfully import!" & vbCrLf _
'        & link & " link data are successfully import!" & vbCrLf _
'        & log_false & " log are unsuccessfully import!", vbInformation, headerMSG
    MsgBox "ALL Data are successfully import!", vbInformation
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    strsql = "Update temp_sal_proses set salary_proses = 0"
    CnG.Execute strsql
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Sub

