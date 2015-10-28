VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_salary_process 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PPH 21 SALARY PROCESS"
   ClientHeight    =   6690
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_trans_salary_process.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_entry 
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   7935
      Begin MSComCtl2.DTPicker DTPicker_month 
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   165871619
         CurrentDate     =   39278
      End
      Begin MSComCtl2.DTPicker DTPicker_periode_from 
         Height          =   300
         Left            =   5040
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   165871619
         CurrentDate     =   39278
      End
      Begin MSComCtl2.DTPicker DTPicker_periode_to 
         Height          =   300
         Left            =   5040
         TabIndex        =   3
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   165871619
         CurrentDate     =   39278
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PERIODE TO"
         Height          =   195
         Left            =   3720
         TabIndex        =   12
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PERIODE FROM"
         Height          =   195
         Left            =   3720
         TabIndex        =   11
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "MONTH"
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   5160
      Width           =   7935
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   4080
         Picture         =   "frm_trans_salary_process.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmd_new 
         Caption         =   "&New"
         Height          =   645
         Left            =   1920
         Picture         =   "frm_trans_salary_process.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin VB.CommandButton cmd_process 
         Caption         =   "&Process"
         Height          =   645
         Left            =   3000
         Picture         =   "frm_trans_salary_process.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   6360
         Picture         =   "frm_trans_salary_process.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8281
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "MONTH"
      Columns(0).DataField=   "month_"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "PERIODE FROM"
      Columns(1).DataField=   "periode_from_"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "PERIODE TO"
      Columns(2).DataField=   "periode_to_"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3942"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3863"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4419"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4339"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=4498"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4419"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
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
      Caption         =   "LIST OF SALARY PROCESSED"
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
      _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
End
Attribute VB_Name = "frm_trans_salary_process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns



Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

'str_sql = "select count(income_code) as rec_count from m_other_income where income_code = '" _
'& Replace$(Trim$(txt_income_code), Chr$(39), Chr$(96)) & "'"
'rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
'
'If rs.Fields("rec_count").Value > 0 Then
'    check_validate_exist_new = True
'    Exit Function
'End If
End Function

Private Sub check_invalid()
MsgBox "Data found!", vbCritical, headerMSG
DTPicker_month.Value = Null
If DTPicker_month.Enabled = True Then DTPicker_month.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If Not DTPicker_month.Value = Adodc1.Recordset.Fields("month").Value And _
check_validate_exist_new Then
    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

'If Trim(txt_income_code) = "" Then
'    MsgBox "Department Code is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_income_code.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
'
'If Trim(txt_income_name) = "" Then
'    MsgBox "Department Name is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_income_name.SetFocus
'    check_validate_new = False
'    Exit Function
'End If


End Function

Private Sub cmd_refresh_Click()
Call load_data
End Sub

Private Sub CmdCancel_Click()
int_mode = 0
Call load_mode
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "No Data selected!", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Are you sure want to delete data '" _
    & Format(TDBGrid1.Columns("month").Value, "mm-yyyy") & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.Execute "delete from h_d_salary where left(month,7) = '" & Format(Adodc1.Recordset.Fields("month").Value, "yyyy-mm") & "'"
CnG.Execute "delete from h_m_salary where left(month,7) = '" & Format(Adodc1.Recordset.Fields("month").Value, "yyyy-mm") & "'"
Call load_data
End Sub

Public Sub set_edit_data()
'With Adodc1.Recordset
'    txt_income_code = .Fields("income_code").Value
'    txt_income_name = .Fields("income_name").Value
'    txt_description = .Fields("description").Value
'End With
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_other_income where income_code = '" _
& Adodc1.Recordset.Fields("income_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub cmd_new_Click()
fra_entry.Visible = True
DTPicker_month = Now
DTPicker_periode_from = Now
DTPicker_periode_to = Now
End Sub

Private Sub cmd_process_Click()
If fra_entry.Visible = False Then
    MsgBox "No data to process!", vbInformation, headerMSG
    Exit Sub
End If

Call process_delete
Call process_procedure
fra_entry.Visible = False
Call load_data
End Sub

Private Sub process_delete()
CnG.Execute "delete from h_d_salary where left(month,7) = '" & Format(DTPicker_month, "yyyy-mm") & "'"
CnG.Execute "delete from h_m_salary where left(month,7) = '" & Format(DTPicker_month, "yyyy-mm") & "'"
End Sub

Private Sub process_procedure()
If fra_entry.Visible = False Then
    Exit Sub
End If

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsx As New ADODB.Recordset
Dim rsy As New ADODB.Recordset
Dim rsz As New ADODB.Recordset

Dim str1 As String
Dim dbl_x As Double
Dim dbl_ot As Double
Dim dbl_z As Double

rs.Open "select * from m_company order by company_code desc", CnG, adOpenStatic, adLockReadOnly
rsx.Open "select * from h_m_salary where month = '2009-07-07'", CnG, adOpenKeyset, adLockOptimistic
rsy.Open "select * from h_d_salary where month = '2009-07-07'", CnG, adOpenKeyset, adLockOptimistic


While Not rs.EOF
    
    str1 = "Call spr_attendance_52_salary_process('" & Format(DTPicker_periode_from, "yyyy-mm-dd") _
    & "', '" & Format(DTPicker_periode_to, "yyyy-mm-dd") & "', 1, '" _
    & rs!COMPANY_CODE & "', 0, '-')"
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open str1, CnG, adOpenStatic, adLockReadOnly
    
    While Not rs1.EOF
        
        With rsy
            .AddNew
            
            .Fields("month").Value = DTPicker_month
            .Fields("periode_from").Value = DTPicker_periode_from
            .Fields("periode_to").Value = DTPicker_periode_to
            .Fields("company_code").Value = rs1.Fields("company_code").Value
            .Fields("company_name").Value = rs1.Fields("company_name").Value
            .Fields("employee_code").Value = rs1.Fields("employee_code").Value
            .Fields("employee_name").Value = rs1.Fields("employee_name").Value
            .Fields("main_salary").Value = rs1.Fields("salary").Value
            
            If rsz.State = 1 Then rsz.Close
            rsz.Open "select * from v_m_salary_row where employee_code = '" _
                & rs1.Fields("employee_code").Value & "'", CnG, adOpenStatic, adLockReadOnly
            If rsz.RecordCount > 0 Then
                dbl_z = rsz!over_time
            Else
                dbl_z = 0
            End If
            
            'MsgBox rs1!total_time_out_ot
            dbl_ot = calc_ot(Right(IIf(IsNull(rs1!total_time_out_ot) = True, "000:00:00", rs1!total_time_out_ot), 9), dbl_z)
            
            dbl_x = rs1.Fields("total_meal_subsidy").Value _
                + rs1.Fields("total_transport_subsidy").Value _
                + rs1.Fields("total_full_presence_subsidy").Value _
                + rs1.Fields("total_phone_voucher_subsidy").Value
            
            .Fields("total_subsidy").Value = dbl_x
            .Fields("total_over_time").Value = dbl_ot
            .Fields("total_reduce").Value = rs1.Fields("salary_reduce").Value
            
            .Fields("net_salary").Value = rs1.Fields("salary").Value _
                + dbl_x - rs1.Fields("salary_reduce").Value
            
            .Update
        End With
        
        rs1.MoveNext
    Wend
    
    rs.MoveNext
Wend

With rsx
    .AddNew
    
    .Fields("month").Value = DTPicker_month
    .Fields("periode_from").Value = DTPicker_periode_from
    .Fields("periode_to").Value = DTPicker_periode_to
    
    .Update
End With
End Sub

Private Function calc_ot(ByVal str_ot As String, ByVal dbl_x As Double) As Double
Dim str1 As String
Dim i As Integer
Dim int_hour, int_minute As Integer


str1 = Trim(str_ot)
If str1 = "" Then str1 = "00:00:00"

i = InStr(1, str1, ":")
int_hour = Int(Left(str1, i - 1))
int_minute = Int(Mid(str1, i + 1, 2))

calc_ot = (int_hour * dbl_x) + ((int_minute / 60) * dbl_x)

'MsgBox int_hour & vbCr & int_minute
End Function

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_other_income where income_code = 'άφ'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 1
Call load_mode
End Sub

Private Sub CmdPrint_Click()
TDBGrid1.PrintInfo.PageSetup
If Not TDBGrid1.PrintInfo.PageSetupCancelled = True Then
    TDBGrid1.PrintInfo.PrintPreview dbgAllRows
End If
End Sub

Private Sub insert_new_data()
'CnG.BeginTrans
'
'With rsBound
'    .AddNew
'
'    .Fields("income_code").Value = Trim(txt_income_code)
'    '-----------------------------------------------------------------------------
'    .Fields("income_name").Value = Trim(txt_income_name)
'    .Fields("description").Value = Trim(txt_description)
'
'    .Update
'End With
'
'CnG.CommitTrans
End Sub

Private Sub edit_old_data()
'On Error GoTo err_capture
'
'CnG.BeginTrans
'With rsBound
'
'    .Fields("income_code").Value = Trim(txt_income_code)
'    '-----------------------------------------------------------------------------
'    .Fields("income_name").Value = Trim(txt_income_name)
'    .Fields("description").Value = Trim(txt_description)
'
'    .Update
'End With
'CnG.CommitTrans
'
'Exit Sub
'err_capture:
'rsBound.CancelBatch adAffectCurrent: rsBound.Close: CnG.RollbackTrans
End Sub

Private Sub CmdSave_Click()
If int_mode = 1 Then
    If Not check_validate_new Then Exit Sub
    If check_validate_exist_new Then
        Call check_invalid: Exit Sub
    End If
    Call insert_new_data
ElseIf int_mode = 2 Then
    If Not check_validate_new Then Exit Sub
    If check_validate_exist_edit Then
        Call check_invalid: Exit Sub
    End If
    Call edit_old_data
End If

Call load_data
int_mode = 0
Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal F As Boolean, ByVal g As Boolean)
'CmdNew.Enabled = a And blnUser_Add
'CmdSave.Enabled = b
'cmdEdit.Enabled = c And blnUser_Edit
'cmdDelete.Enabled = d And blnUser_Delete
'CmdCancel.Enabled = e
'
'CmdPrint.Enabled = f
'cmd_refresh.Enabled = g
End Sub

Private Sub clear_view_data()
Dim Ctr As Control
For Each Ctr In Me
    If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
        If Not LCase(Ctr.name) = "txt_company_name" Then Ctr.Text = ""
    ElseIf TypeOf Ctr Is TDBCombo Then
        If Not LCase(Ctr.name) = "tdbcombo_company" Then Ctr.Text = ""
    ElseIf TypeOf Ctr Is DTPicker Then
        Ctr.Value = Now
    End If
Next
End Sub

Private Sub set_enabled_control(ByVal i As Boolean)
Dim Ctr As Control
For Each Ctr In Me
    If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
        Ctr.Enabled = i
    ElseIf TypeOf Ctr Is TDBCombo Then
        Ctr.Enabled = i
    ElseIf TypeOf Ctr Is DTPicker Then
        Ctr.Value = Now
        Ctr.Enabled = i
    End If
Next
End Sub

Private Sub set_new_data()
'cbo_jns_kelamin.ListIndex = 1
End Sub

Private Sub set_data_mode()
'If int_mode = 1 Then        'NEW
'    Call clear_view_data
'    fra_entry.Visible = True
'    txt_income_code.Enabled = True
'    TDBGrid1.Enabled = False
'    Call set_new_data
'
'    If txt_income_code.Enabled = True Then
'        txt_income_code.SetFocus
'    End If
'
'ElseIf int_mode = 0 Then    'VIEW
'    Call clear_view_data
'    fra_entry.Visible = False
'    TDBGrid1.Enabled = True
'
'ElseIf int_mode = 2 Then    'EDIT
'    Call set_edit_data
'    txt_income_code.Enabled = False
'    fra_entry.Visible = True
'    TDBGrid1.Enabled = False
'End If
End Sub

Private Sub load_mode()
If int_mode = 1 Then        ' new
    Call set_buttons_enable(False, True, False, False, True, False, False)
ElseIf int_mode = 0 Then    ' view
    Call set_buttons_enable(True, False, True, True, False, True, True)
ElseIf int_mode = 2 Then    ' edit/revise
    Call set_buttons_enable(False, True, False, False, True, False, False)
End If

Call set_data_mode
End Sub

Private Sub Command1_Click()
'Call calc_ot("115:45:55")
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = strConn

Call load_data

Call load_data_user_access(Me)
int_mode = 0
Call load_mode
Timer1.Enabled = True
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select

End Sub

Private Sub txtTelp1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtTelp2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Function get_inc_kode() As String
Dim rs As New ADODB.Recordset
Dim str_inc_kode As String

rs.Open "select max(kode_supplier) as curr_kode from m_supplier", CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    If IsNull(rs.Fields("curr_kode").Value) = True Then
        str_inc_kode = "00001"
    Else
        str_inc_kode = rs.Fields("curr_kode").Value
        str_inc_kode = Right("0000" & Trim(str(CLng(str_inc_kode) + 1)), 5)
    End If
End If

get_inc_kode = str_inc_kode
End Function


Private Sub txtTelp3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub clear_filter()
For Each Col In TDBGrid1.Columns
    Col.FilterText = ""
Next Col
Adodc1.Recordset.Filter = adFilterNone
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
        
        tmp = tmp & Col.DataField & " LIKE '" & Col.FilterText & "*'"
    End If
Next Col
getFilter = tmp
End Function

Private Sub TDBGrid1_FilterChange()
On Error GoTo ErrHandler

Dim i As Integer

Set Cols = TDBGrid1.Columns
i = TDBGrid1.Col
TDBGrid1.HoldFields

Adodc1.Recordset.Filter = getFilter()
TDBGrid1.Col = i
TDBGrid1.EditActive = True

TDBGrid1.SelStart = Len(TDBGrid1.Columns(i).FilterText)
If TDBGrid1.ApproxCount < 1 Then
    Call clear_filter
    TDBGrid1.Col = i
End If

Exit Sub
ErrHandler:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data()
Adodc1.RecordSource = "select *, cast(left(month,7) as char) as month_, cast(left(periode_from,10) as char) as periode_from_, cast(left(periode_to,10) as char) as periode_to_ " _
& "from h_m_salary order by month asc"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
End Sub




