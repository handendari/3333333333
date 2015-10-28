VERSION 5.00
Object = "{D28F8786-0BB9-402B-92DC-F32DE23A324E}#3.0#0"; "OutlookBar.ocx"
Object = "{5B033ECF-098E-11D1-A4B2-444553540000}#1.0#0"; "Subclass.ocx"
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.1#0"; "cPopMenu6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdi_absensi 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000A&
   Caption         =   "FEI ATTENDANCE & PAYROLL SYSTEM 2.0"
   ClientHeight    =   8535
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "mdi_absensi.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "mdi_absensi.frx":058A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Height          =   8190
      Left            =   7635
      ScaleHeight     =   8130
      ScaleWidth      =   4065
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4125
      Begin VB.TextBox txt_hwnd 
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Left            =   2280
         ScaleHeight     =   435
         ScaleWidth      =   315
         TabIndex        =   5
         Top             =   5760
         Width           =   375
      End
      Begin SubclassCtl.Subclass Subclass1 
         Left            =   1440
         Top             =   5520
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   6120
         Width           =   975
      End
      Begin VB.Timer timer_terminate 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   600
         Top             =   4800
      End
      Begin cPopMenu6.PopMenu ctlPopMenu 
         Left            =   1440
         Top             =   6360
         _ExtentX        =   1058
         _ExtentY        =   1058
         HighlightCheckedItems=   0   'False
         TickIconIndex   =   0
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   600
         Top             =   5280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdi_absensi.frx":258F7
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdi_absensi.frx":25E91
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdi_absensi.frx":262E3
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image img_false 
         Height          =   960
         Left            =   0
         Picture         =   "mdi_absensi.frx":26635
         Top             =   2880
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image img_true 
         Height          =   960
         Left            =   0
         Picture         =   "mdi_absensi.frx":2737F
         Top             =   4080
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8190
      Left            =   2535
      ScaleHeight     =   8190
      ScaleWidth      =   120
      TabIndex        =   2
      Top             =   0
      Width           =   125
      Begin prj_genting.vbButton cmd_navigation 
         Height          =   2055
         Left            =   0
         TabIndex        =   7
         Top             =   2640
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   3625
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "mdi_absensi.frx":280C9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Timer timer_free_trial_30 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5280
      Top             =   360
   End
   Begin VB.Timer timer_get_log_data 
      Interval        =   60000
      Left            =   4800
      Top             =   360
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   8190
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2647
            MinWidth        =   2647
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6704
            MinWidth        =   6704
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "mdi_absensi.frx":280E5
         EndProperty
      EndProperty
   End
   Begin OutlookBar.ctxOutlookBar ctxOutlookBar1 
      Align           =   3  'Align Left
      Height          =   8190
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   14446
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormatControl   =   "mdi_absensi.frx":2823F
      FormatGroup     =   "mdi_absensi.frx":283B3
      FormatGroupHover=   "mdi_absensi.frx":284D7
      FormatGroupPressed=   "mdi_absensi.frx":285E7
      FormatGroupSelected=   "mdi_absensi.frx":286EB
      FormatItem      =   "mdi_absensi.frx":287D3
      FormatItemLargeIcons=   "mdi_absensi.frx":288E3
      FormatItemHover =   "mdi_absensi.frx":289DF
      FormatItemPressed=   "mdi_absensi.frx":28AC7
      FormatItemSelected=   "mdi_absensi.frx":28BAF
      FormatSmallIcon =   "mdi_absensi.frx":28C6F
      FormatSmallIconHover=   "mdi_absensi.frx":28D57
      FormatSmallIconPressed=   "mdi_absensi.frx":28E53
      FormatSmallIconSelected=   "mdi_absensi.frx":28F4F
      FormatLargeIcon =   "mdi_absensi.frx":2904B
      FormatLargeIconHover=   "mdi_absensi.frx":29133
      FormatLargeIconPressed=   "mdi_absensi.frx":2922F
      FormatLargeIconSelected=   "mdi_absensi.frx":2932B
      Groups          =   "mdi_absensi.frx":29427
   End
   Begin VB.Menu M01 
      Caption         =   "A1"
   End
   Begin VB.Menu M02 
      Caption         =   "2"
   End
   Begin VB.Menu M03 
      Caption         =   "3"
   End
   Begin VB.Menu M04 
      Caption         =   "4"
   End
   Begin VB.Menu M05 
      Caption         =   "5"
   End
   Begin VB.Menu M06 
      Caption         =   "6"
   End
   Begin VB.Menu M07 
      Caption         =   "7"
   End
   Begin VB.Menu M08 
      Caption         =   "8"
   End
   Begin VB.Menu M09 
      Caption         =   "9"
   End
   Begin VB.Menu M10 
      Caption         =   "10"
   End
End
Attribute VB_Name = "mdi_absensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit


Dim str_form_title(60) As String, int_count_menu As Integer
Dim mnu As Menu
Dim bln_navigation As Boolean


'Standard rectangle structure
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'Win API declarations
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Win API constants
Private Const SRCCOPY = &HCC0020
Private Const GW_CHILD = 5
Private Const COLOR_APPWORKSPACE = 12
Private Const WM_ERASEBKGND = &H14
Private Const WM_PAINT = &HF


Private Declare Function EbExecuteLine Lib "vba6.dll" _
        (ByVal pStringToExec As Long, ByVal Foo1 As Long, _
        ByVal Foo2 As Long, ByVal fCheckOnly As Long) As Long

'-- menu
Private Declare Function TrackPopupMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private m_lAboutId As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long


Function FExecuteCode(stCode As String, _
            Optional fCheckOnly As Boolean) As Boolean

    FExecuteCode = EbExecuteLine(StrPtr(stCode), 0&, 0&, Abs(fCheckOnly)) = 0
End Function

Private Sub set_menu_handle()
Dim l As Long
Dim lIndex As Long
Dim lC As Long

    With ctlPopMenu
        
        ' Remove the close item from the system menu:
        lC = .SystemMenuCount
        .SystemMenuRemoveItem lC
        ' Add a new about item to the system menu:
        'm_lAboutId = .SystemMenuAppendItem("&About...")
        .OfficeXpStyle = True
        
        
        ' Associate the image list:
        '.ImageList = ilsIcons
        
        ' Parse through the VB designed menu and sub class the items:
        .SubClassMenu Me
        
'        lIndex = .MenuIndex("mnuEdit(2)")
'        chkEnable.value = .Enabled(lIndex) * -1
'
'        ' Add the icons:
'        pSetIcon "OPEN", "mnuFile(0)"
'        pSetIcon "SAVE", "mnuFile(1)"
'        pSetIcon "PRINT", "mnuFile(3)"
'
'        pSetIcon "CUT", "mnuEdit(0)"
'        pSetIcon "COPY", "mnuEdit(1)"
'        pSetIcon "PASTE", "mnuEdit(2)"
'        ilsIcons.ListImages(ctlPopMenu.ItemIcon("mnuEdit(2)") + 1).Draw picIcon.hdc, 4 * Screen.TwipsPerPixelX, 0.4 * Screen.TwipsPerPixelY, imlTransparent
'        picIcon.Refresh
'        pSetIcon "FIND", "mnuEdit(4)"
'
'        pSetIcon "HELP", "mnuHelp(0)"
'        pSetIcon "NET", "mnuHelp(1)"
'
'        .TickIconIndex = plGetIconIndex("TICK")
    End With
    
    ' Add a whole new set of menu items and sub items to the last
    ' menu item:
    'pCreateMenuItems
        
End Sub

Private Sub create_v_menu()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim b As PictureBox
Dim bln_menu As Boolean

'Set b = Picture2
'Set b.Picture = LoadPicture(App.Path & "\icon\box2.ico")

rs1.Open "select * from m_menu order by menu_code asc", CnG, adOpenStatic, adLockReadOnly
If rs1.RecordCount > 0 Then rs1.MoveFirst

ctxOutlookBar1.Groups.Clear
While Not rs1.EOF
    ctxOutlookBar1.Groups.Add UCase(rs1.Fields("menu_name").Value), , , rs1.Fields("menu_code").Value
    rs1.MoveNext
Wend

rs2.Open "select * from m_sub_menu where flag_detail=1 order by sub_menu_code asc", CnG, adOpenStatic, adLockReadOnly
If rs2.RecordCount > 0 Then rs2.MoveFirst

While Not rs2.EOF
    ctxOutlookBar1.Groups.item(rs2.Fields("menu_code").Value).GroupItems.Add rs2.Fields("sub_menu_name").Value, _
        , , rs2.Fields("sub_menu_code").Value
            
    If Val("" & rs2.Fields("flag_detail").Value) = 0 Or _
    (rs2.Fields("flag_detail").Value = 1 And rs2.Fields("form_name").Value = "") Then
        bln_menu = True
    Else
        bln_menu = get_menu_user_access(rs2.Fields("form_name").Value)
    End If
    
    ctxOutlookBar1.Groups.item(rs2.Fields("menu_code").Value).GroupItems.item( _
        rs2.Fields("sub_menu_code")).Enabled = bln_menu
            
    rs2.MoveNext
Wend
End Sub

Private Sub hide_h_menu()
Dim rs1 As New ADODB.Recordset
Dim str1 As String
Dim lng_menu As Long, i As Integer

str1 = "select * from m_menu_constant"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly

If rs1.RecordCount > 0 Then
    rs1.MoveFirst
    While Not rs1.EOF
        Set mnu = Controls(rs1.Fields("menu_code").Value)
        mnu.Visible = False

        rs1.MoveNext
        i = i + 1
    Wend
End If
End Sub

Private Sub unhide_h_menu()
Dim rs1 As New ADODB.Recordset
Dim str1 As String
Dim lng_menu As Long, i As Integer

str1 = "select * from m_menu_constant"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly

If rs1.RecordCount > 0 Then
    rs1.MoveFirst
    While Not rs1.EOF
        Set mnu = Controls(rs1.Fields("menu_code").Value)
        mnu.Visible = True

        rs1.MoveNext
        i = i + 1
    Wend
End If


End Sub

Private Sub hide_unreg_menu()
Dim rs1 As New ADODB.Recordset
Dim str1 As String
Dim lng_menu As Long, i As Integer

str1 = "select * from m_menu_constant where menu_code not in (select menu_code from m_menu)"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly

If rs1.RecordCount > 0 Then
    rs1.MoveFirst
    While Not rs1.EOF
        Set mnu = Controls(rs1.Fields("menu_code").Value)
        mnu.Visible = False

        rs1.MoveNext
        i = i + 1
    Wend
End If
End Sub

Private Sub set_menu_caption()
Dim rs1 As New ADODB.Recordset
Dim str1 As String
Dim lng_menu As Long, i As Integer

str1 = "select * from m_menu order by menu_code"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly

If rs1.RecordCount > 0 Then
    rs1.MoveFirst
    While Not rs1.EOF
        Set mnu = Controls(rs1.Fields("menu_code").Value)
        mnu.Caption = rs1.Fields("menu_name").Value

        rs1.MoveNext
        i = i + 1
    Wend
End If
End Sub

Private Sub create_h_menu()
Dim rs1 As New ADODB.Recordset
Dim str1 As String
Dim lng_menu As Long, i, j As Integer
Dim bln_menu As Boolean


str1 = "select max(level) as rec from m_sub_menu"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly
If Not rs1.RecordCount > 0 Then Exit Sub
j = Val(rs1.Fields("rec"))

For i = 2 To j

If rs1.State = 1 Then rs1.Close
str1 = "select * from m_sub_menu where level=" & i & " order by sub_menu_code"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly

If rs1.RecordCount > 0 Then
    rs1.MoveFirst
    While Not rs1.EOF
        With ctlPopMenu
            If Not (.MenuExists(rs1.Fields("sub_menu_code").Value)) Then
                If i = 2 Then
                    lng_menu = .MenuIndex(rs1.Fields("menu_code").Value)
                Else
                    lng_menu = .MenuIndex(rs1.Fields("parent_menu_code").Value)
                End If
                
                If rs1.Fields("sub_menu_name").Value = "-" Then
                    .AddItem "-", rs1.Fields("sub_menu_code").Value, , , lng_menu
                Else
                    If Val("" & rs1.Fields("flag_detail").Value) = 0 Or _
                    (rs1.Fields("flag_detail").Value = 1 And rs1.Fields("form_name").Value = "") Then
                        bln_menu = True
                    Else
                        bln_menu = get_menu_user_access(rs1.Fields("form_name").Value)
                    End If
                    
                    .AddItem rs1.Fields("sub_menu_name").Value, rs1.Fields("sub_menu_code").Value, , , lng_menu _
                        , 2, , bln_menu
                        
                    
                End If
               
            Else
               'MsgBox "Menu items are already added.", vbInformation
            End If
        End With

        rs1.MoveNext
    Wend
End If

Next i
End Sub

Private Sub remove_h_menu()
Dim rs1 As New ADODB.Recordset
Dim str1 As String
Dim lng_menu As Long, i, j As Integer
Dim bln_menu As Boolean


str1 = "select max(level) as rec from m_sub_menu"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly
If Not rs1.RecordCount > 0 Then Exit Sub
j = Val(rs1.Fields("rec"))

For i = j To 2 Step -1

If rs1.State = 1 Then rs1.Close
str1 = "select * from m_sub_menu where level=" & i & " order by sub_menu_code"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly

If rs1.RecordCount > 0 Then
    rs1.MoveFirst
    While Not rs1.EOF
        With ctlPopMenu
            If .MenuExists(rs1.Fields("sub_menu_code").Value) Then
                .RemoveItem rs1.Fields("sub_menu_code").Value
               
            Else
               'MsgBox "Menu items are already added.", vbInformation
            End If
        End With

        rs1.MoveNext
    Wend
End If

Next i
End Sub

Private Function get_menu_user_access(ByVal str_frm As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String

If LOGIN_LEVEL = 100 Then
    blnUser_Read = True
    blnUser_Add = True
    blnUser_Edit = True
    blnUser_Delete = True
    blnUser_Posting = True
    blnUser_Printing = True
    
    get_menu_user_access = blnUser_Read
    Exit Function
End If

blnUser_Read = False
blnUser_Add = False
blnUser_Edit = False
blnUser_Delete = False
blnUser_Posting = False
blnUser_Printing = False

str_sql = "select F.sub_menu_code, F.form_name, F.form_title, " _
& "F.allow_read, F.allow_add, F.allow_edit, F.allow_delete, F.allow_post, F.allow_print " _
& "from m_user U join m_employee b on U.employee_code = b.employee_code " _
& "join t_user F on F.level_code = U.user_code " _
& "Where U.user_code = '" & LOGIN_CODE & "' and user_name = '" & LOGIN_NAME _
& "' and user_pass = '" & LOGIN_PASS & "' and upper(F.form_name)='" & UCase(str_frm) & "'"

rs.Open str_sql, CnG, adOpenStatic, adLockBatchOptimistic
If rs.RecordCount > 0 Then

    'If UCase(rs.Fields("form_title").value) = UCase(str_FormCaption) Then
        
        blnUser_Read = rs.Fields("allow_read").Value
        blnUser_Add = rs.Fields("allow_add").Value
        blnUser_Edit = rs.Fields("allow_edit").Value
        blnUser_Delete = rs.Fields("allow_delete").Value
        blnUser_Posting = rs.Fields("allow_post").Value
        blnUser_Printing = rs.Fields("allow_print").Value
        
        'rs.MoveLast
    'End If

End If

get_menu_user_access = blnUser_Read
End Function

Private Sub cmd_navigation_Click()
bln_navigation = Not bln_navigation
Call set_navigation
Call MDIForm_Resize
End Sub

Private Sub ctlPopMenu_Click(ItemNumber As Long)
Call exec_menu(ctlPopMenu.MenuKey(ItemNumber))
End Sub

Private Sub ctxOutlookBar1_ButtonClick(ByVal oBtn As OutlookBar.cButton)
Dim rs1 As New ADODB.Recordset

If Trim(oBtn.Key) <> "" Then

Call exec_menu(oBtn.Key)
'    If oBtn.Caption = "Cascade" Then
'        mdi_accounting.Arrange vbCascade
'        Exit Sub
'    ElseIf oBtn.Caption = "Tile Horizontal" Then
'        mdi_accounting.Arrange vbTileHorizontal
'        Exit Sub
'    ElseIf oBtn.Caption = "Tile Vertical" Then
'        mdi_accounting.Arrange vbTileVertical
'        Exit Sub
'    ElseIf oBtn.Caption = "Exit" Then
'        End
'    End If

'    rs1.Open "select * from m_sub_menu where sub_menu_code='" & oBtn.Key & "'", cng, adOpenStatic, adLockReadOnly
'    If rs1.RecordCount > 0 Then
'        Call exec_menu(oBtn.Key)
'    End If
End If
End Sub

Private Sub exec_menu_backup(ByVal str_key As String)
On Error GoTo err_handle
Dim rs1 As New ADODB.Recordset
Dim str1 As String

str1 = "select * from m_sub_menu where sub_menu_code = '" & str_key & "'"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly
If Not rs1.RecordCount > 0 Then Exit Sub

If UCase(rs1.Fields("sub_menu_name").Value) = "EXIT" Or UCase(rs1.Fields("sub_menu_name").Value) = "KELUAR" Then
    'cmd_navigation.end_vbbutton
    'Unload Me
    'ctlPopMenu.UnsubclassMenu
    'End
    timer_terminate.Enabled = True
End If


'If rs1.Fields("form_name").value = "frm_etc_login" Then
'    Call exec_menu_(frm_etc_login, rs1.Fields("form_modal").value)
'ElseIf rs1.Fields("form_name").value = "frm_etc_about" Then
'    Call exec_menu_(frm_etc_about, rs1.Fields("form_modal").value)
'    ' ...
'End If

'++++++++++++++++++++++FILE++++++++++++++++
If rs1.Fields("form_name").Value = "frm_stg_all" Then Call exec_menu_(frm_stg_all, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_change_password" Then Call exec_menu_(frm_change_password, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_etc_login" Then Call exec_menu_(frm_etc_login, rs1.Fields("form_modal").Value)
'++++++++++++++++++++++++++++++++++++++++++

'********************MASTER****************
If rs1.Fields("form_name").Value = "frm_mst_form" Then Call exec_menu_(frm_mst_form, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_menu" Then Call exec_menu_(frm_mst_menu, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_group_hak_akses" Then Call exec_menu_(frm_mst_group_hak_akses, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_user" Then Call exec_menu_(frm_mst_user, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_device" Then Call exec_menu_(frm_mst_device, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_enroll" Then Call exec_menu_(frm_mst_enroll, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_enroll_uareu" Then Call exec_menu_(frm_mst_enroll_uareu, rs1.Fields("form_modal").Value)
'------------
If rs1.Fields("form_name").Value = "frm_mst_bank" Then Call exec_menu_(frm_mst_bank, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_company" Then Call exec_menu_(frm_mst_company, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_department" Then Call exec_menu_(frm_mst_department, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_division" Then Call exec_menu_(frm_mst_division, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_location" Then Call exec_menu_(frm_mst_location, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_title" Then Call exec_menu_(frm_mst_title, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_education" Then Call exec_menu_(frm_mst_education, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_family_level" Then Call exec_menu_(frm_mst_family_level, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_family_from" Then Call exec_menu_(frm_mst_family_from, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_education_history" Then Call exec_menu_(frm_mst_education_history, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_education_non_history" Then Call exec_menu_(frm_mst_education_non_history, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_job_history" Then Call exec_menu_(frm_mst_job_history, rs1.Fields("form_modal").Value)
'-----------
If rs1.Fields("form_name").Value = "frm_mst_employee" Then Call exec_menu_(frm_mst_employee, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_user_level" Then Call exec_menu_(frm_mst_user_level, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_working_time" Then Call exec_menu_(frm_mst_working_time, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_working_day" Then Call exec_menu_(frm_mst_working_day, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_employees_working_time" Then Call exec_menu_(frm_mst_employees_working_time, rs1.Fields("form_modal").Value)

If rs1.Fields("form_name").Value = "frm_mst_prestasi" Then Call exec_menu_(frm_mst_prestasi, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_salary_standard" Then Call exec_menu_(frm_mst_salary_standard, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_employee_level" Then Call exec_menu_(frm_mst_group_hak_akses, rs1.Fields("form_modal").Value)

If rs1.Fields("form_name").Value = "frm_mst_other_income" Then Call exec_menu_(frm_mst_other_income, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_other_expense" Then Call exec_menu_(frm_mst_other_expense, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_salary_item" Then Call exec_menu_(frm_mst_salary_item, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_salary_summary" Then Call exec_menu_(frm_mst_salary_summary, rs1.Fields("form_modal").Value)

If rs1.Fields("form_name").Value = "frm_mst_salary_inc_employee" Then Call exec_menu_(frm_mst_salary_inc_employee, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_salary_exp_employee" Then Call exec_menu_(frm_mst_salary_exp_employee, rs1.Fields("form_modal").Value)

If rs1.Fields("form_name").Value = "frm_mst_ptkp" Then Call exec_menu_(frm_mst_ptkp, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_pph21" Then Call exec_menu_(frm_mst_pph21, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_jamsostek" Then Call exec_menu_(frm_mst_jamsostek, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_status" Then Call exec_menu_(frm_mst_status, rs1.Fields("form_modal").Value)
'******************************************

'&&&&&&&&&&&&&&&&UTILITY&&&&&&&&&&&&&&&&&&
If rs1.Fields("form_name").Value = "frm_trans_log_attendance" Then Call exec_menu_(frm_trans_log_attendance, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_import_log_attendance" Then Call exec_menu_(frm_trans_import_log_attendance, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_log_attendance_uareu" Then Call exec_menu_(frm_trans_log_attendance_uareu, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_import_attendance" Then Call exec_menu_(frm_trans_import_attendance, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_manual_check" Then Call exec_menu_(frm_trans_manual_check, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_list_manual_overtime" Then Call exec_menu_(frm_list_manual_overtime, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_absent" Then Call exec_menu_(frm_trans_absent, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_list_manual_att" Then Call exec_menu_(frm_list_manual_att, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_leave" Then Call exec_menu_(frm_trans_leave, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_general_leave" Then Call exec_menu_(frm_trans_general_leave, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_summary_leave" Then Call exec_menu_(frm_trans_summary_leave, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_duty" Then Call exec_menu_(frm_trans_duty, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_holiday" Then Call exec_menu_(frm_trans_holiday, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_performance" Then Call exec_menu_(frm_trans_performance, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_salary_process" Then Call exec_menu_(frm_trans_salary_process, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_loan" Then Call exec_menu_(frm_trans_loan, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_proses_thr" Then Call exec_menu_(frm_proses_thr, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_history_prestasi" Then Call exec_menu_(frm_mst_history_prestasi, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_tranfer_data" Then Call exec_menu_(frm_trans_tranfer_data, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_mst_salary_exp_employee_installment" Then Call exec_menu_(frm_mst_salary_exp_employee_installment, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_List_pensiun" Then Call exec_menu_(frm_List_pensiun, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_export_pph" Then Call exec_menu_(frm_export_pph, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_List_IncomeLain" Then Call exec_menu_(frm_List_IncomeLain, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_List_PotonganLain" Then Call exec_menu_(frm_List_PotonganLain, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_trans_bpjs" Then Call exec_menu_(frm_trans_bpjs, rs1.Fields("form_modal").Value)
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

'#############REPORT######################
If rs1.Fields("form_name").Value = "frm_rpt_master" Then Call exec_menu_(frm_rpt_master, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_summary_employee" Then Call exec_menu_(frm_rpt_summary_employee, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_biodata" Then Call exec_menu_(frm_rpt_biodata, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_performance" Then Call exec_menu_(frm_rpt_performance, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_detail_attendance" Then Call exec_menu_(frm_rpt_detail_attendance, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_summary_attendance" Then Call exec_menu_(frm_rpt_summary_attendance, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_verify_att" Then Call exec_menu_(frm_rpt_verify_att, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_summary_leave" Then Call exec_menu_(frm_rpt_summary_leave, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_summary_salary" Then Call exec_menu_(frm_rpt_summary_salary, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_summary_salary_ketapang" Then Call exec_menu_(frm_rpt_summary_salary_ketapang, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_overtime" Then Call exec_menu_(frm_rpt_overtime, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_salary_bank" Then Call exec_menu_(frm_rpt_salary_bank, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_summary_pph21" Then Call exec_menu_(frm_rpt_summary_pph21, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_log_mail" Then Call exec_menu_(frm_rpt_log_mail, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_log_activity" Then Call exec_menu_(frm_rpt_log_activity, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_premi_jamsostek" Then Call exec_menu_(frm_rpt_premi_jamsostek, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_loan" Then Call exec_menu_(frm_rpt_loan, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_man_power" Then Call exec_menu_(frm_rpt_man_power, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_rpt_layout_absen" Then Call exec_menu_(frm_rpt_layout_absen, rs1.Fields("form_modal").Value)

If rs1.Fields("form_name").Value = "frm_etc_register" Then Call exec_menu_(frm_etc_register, rs1.Fields("form_modal").Value)
If rs1.Fields("form_name").Value = "frm_etc_about" Then Call exec_menu_(frm_etc_about, rs1.Fields("form_modal").Value)
'#########################################

'If rs1.Fields("form_name").Value = "frm_rpt_summary_salary_est" Then Call exec_menu_(frm_rpt_summary_salary_est, rs1.Fields("form_modal").Value)
'If rs1.Fields("form_name").Value = "frm_trans_import_attendance" Then Call exec_menu_(frm_trans_import_attendance, rs1.Fields("form_modal").Value)

'If rs1.Fields("form_name").Value = "frm_mst_salary_slip" Then Call exec_menu_(frm_mst_salary_slip, rs1.Fields("form_modal").Value)
'If rs1.Fields("form_name").Value = "frm_trans_spl" Then Call exec_menu_(frm_trans_spl, rs1.Fields("form_modal").Value)
'If rs1.Fields("form_name").Value = "frm_rpt_summary_p24" Then Call exec_menu_(frm_rpt_summary_p24, rs1.Fields("form_modal").Value)
'If rs1.Fields("form_name").Value = "frm_rpt_spt_masa" Then Call exec_menu_(frm_rpt_spt_masa, rs1.Fields("form_modal").Value)



    '++++++++++INSERT UNTUK LOG AKTIFITAS USER++++++++++++
    Dim clsFn As New clsFunction
    clsFn.InsertLog (rs1!form_title)
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++
    
Exit Sub
err_handle:
MsgBox Err.Description
End Sub

Private Sub exec_menu_(ByRef frm As Form, ByVal i As Integer)
frm.Top = 1
frm.Left = 1
frm.Show i
Call set_sizeable_form(frm)
End Sub

Private Sub exec_menu(ByVal str_key As String)
'On Error GoTo err_handle
Dim rs1 As New ADODB.Recordset
Dim str1, str2, stre As String


Call exec_menu_backup(str_key)
'Exit Sub

'
'
'    str1 = "select * from m_sub_menu where sub_menu_code = '" & str_key & "'"
'    rs1.Open str1, CnG, adOpenStatic, adLockReadOnly
'    If Not rs1.RecordCount > 0 Then Exit Sub
'
'    If UCase(rs1.Fields("sub_menu_name").Value) = "EXIT" Then
'        timer_terminate.Enabled = True
'        Exit Sub
'    End If
'
'    str1 = rs1.Fields("form_name").Value
'    str2 = rs1.Fields("form_modal").Value
'    stre = str1 & ".left=1:" & str1 & ".top=1:" & str1 & ".show " & str2
'    Call FExecuteCode(stre)
'    Call set_sizeable_form_by_hwnd(str1)
'
'    '++++++++++INSERT UNTUK LOG AKTIFITAS USER++++++++++++
'    Dim clsFn As New clsFunction
'    clsFn.InsertLog ("Click Menu : " & rs1!form_title)
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++
    
'Exit Sub
'err_handle:
'MsgBox "Form " & str1 & "Tidak Ada Pada Program...!!"
End Sub

Private Function isFormExist(str1 As String) As Boolean
Dim objForm As Form
    isFormExist = False
    For Each objForm In VB.App
        If (Trim(objForm.name) = Trim(str1)) Then
            isFormExist = True
        End If
    Next
    'MsgBox "Load Status: " & FlgLoaded & vbCrLf & "Show Status:" & FlgShown
End Function
    
Private Sub MDIForm_Load()
With StatusBar1
    .Panels(1).Text = " DATE : " & UCase(Format(Date, "dddd, dd MMMM yyyy"))
    .Panels(4).Text = "LOGIN NAME  : " & UCase(LOGIN_NAME)
End With

'Call load_all_menu
'Call disable_all_menu
'Call load_user_menu

BLN_RUNNING = True
MDI_STATE = True

Call get_setting_device
Call get_setting_auto_log
timer_get_log_data.Enabled = BLN_AUTO_LOG

Call show_app_title

ctlPopMenu.ImageList = ImageList1
Call set_menu_handle
Call re_create_menu

bln_navigation = True
Call set_navigation
End Sub

Public Sub re_create_menu()
Call delete_sub_menu
ctxOutlookBar1.Groups.Clear
Call hide_unreg_menu
Call set_menu_caption

Call create_h_menu
Call create_v_menu

With StatusBar1
    .Panels(1).Text = " DATE : " & UCase(Format(Date, "dddd, dd MMMM yyyy"))
    .Panels(4).Text = "LOGIN NAME  : " & UCase(LOGIN_NAME)
End With
End Sub

Private Sub delete_sub_menu()
Dim rs1 As New ADODB.Recordset
Dim str1 As String
Dim lng_menu As Long, i As Integer

str1 = "select * from m_menu_constant order by menu_code asc"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly

If rs1.RecordCount > 0 Then
    rs1.MoveFirst
    While Not rs1.EOF
        Set mnu = Controls(rs1.Fields("menu_code").Value)
        ctlPopMenu.ClearSubMenusOfItem (rs1.Fields("menu_code").Value)

        rs1.MoveNext
        i = i + 1
    Wend
End If
End Sub

Private Sub set_navigation()
ctxOutlookBar1.Visible = bln_navigation
Set cmd_navigation.PictureNormal = IIf(bln_navigation, img_true, img_false)
Set cmd_navigation.PictureOver = IIf(bln_navigation, img_true, img_false)

If bln_navigation = True Then
    Call remove_h_menu
    Call hide_h_menu
Else
    Call unhide_h_menu
    Call hide_unreg_menu
    Call create_h_menu
End If
End Sub

Public Sub show_app_title()
If IS_LICENSED Then
    Me.Caption = "ATTENDANCE & PAYROLL SYSTEM 2.0 (LICENSED)"
Else
    Me.Caption = "ATTENDANCE & PAYROLL SYSTEM 2.0 (NO LICENSE)"
End If
End Sub


Private Sub MDIForm_Resize()
Dim i As Integer
 
i = (Me.ScaleHeight - cmd_navigation.Height) / 2
cmd_navigation.Top = i
End Sub

Private Sub timer_get_log_data_Timer()
int_timer_tick = int_timer_tick + 1

If check_auto_log Then
    Dim rs As New ADODB.Recordset
    
    'Call mnu_trans_log_att_Click
    frm_trans_log_attendance.public_int_caller = 1
    frm_trans_log_attendance.Show
    
    
    rs.Open "select * from m_device order by ip_address asc", CnG, adOpenStatic, adLockReadOnly
    If Not rs.RecordCount > 0 Then
        Unload frm_trans_log_attendance
        Exit Sub
    End If
    
    rs.MoveFirst
    While Not rs.EOF
        FG_IP_ADDRESS = rs.Fields("ip_address").Value
        FG_PORT_NUMBER = rs.Fields("port_number").Value
        Call frm_trans_log_attendance.download_action
        
        rs.MoveNext
    Wend
    
    Unload frm_trans_log_attendance
    int_timer_tick = 0
End If
End Sub

Private Function check_auto_log() As Boolean

Exit Function

Dim rs As New ADODB.Recordset
check_auto_log = False

rs.Open "select * from s_device where s_number = 1", CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount = 1 Then
    If rs.Fields("s_device").Value = 1 Then
        check_auto_log = True
    End If
Else
    check_auto_log = False
    Exit Function
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from s_download_mode where s_number = 1", CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount = 1 Then
    If rs.Fields("s_auto_download").Value = 1 Then
        check_auto_log = True
    End If
Else
    check_auto_log = False
    Exit Function
End If


If rs.State = 1 Then rs.Close
rs.Open "select count(*) as rec_count from s_auto_log where left(cast(s_time as time),5)='" _
    & Format(Now, "hh:mm") & "' and ifnull(s_enable,0)=1", CnG, adOpenStatic, adLockReadOnly
'check_auto_log = IIf(rs.Fields("rec_count").Value >= 1 And mnu_trans_log_att.Enabled = True, True, False)
End Function

Private Sub timer_free_trial_30_Timer()
Call check_free_trial_30
End Sub

Public Sub check_free_trial_30()
Dim rs As New ADODB.Recordset
rs.Open "select sysdate() as dt", CnG, adOpenStatic, adLockReadOnly

If rs.RecordCount > 0 Then
    If Format(rs!Dt, "yyyy-mm-dd") >= "2009-11-01" Then
        MsgBox "30 days to trial was expired..." & vbCr & "Please Order The License!", vbCritical, headerMSG
        End
    End If
Else
    End
End If

End Sub

Private Sub load_user_menu()
Dim rs As New ADODB.Recordset
Dim str_sql As String

If LOGIN_LEVEL = 100 Then
    Call enable_all_menu
    Exit Sub
End If

str_sql = "select F.form_code, F.form_name, F.form_title, " _
& "F.allow_read, F.allow_add, F.allow_edit, F.allow_delete " _
& "from m_user U, t_user F " _
& "Where U.user_code = " & LOGIN_CODE & " and user_name = '" & LOGIN_NAME _
& "' and user_pass = '" & LOGIN_PASS & "' and U.user_code=F.user_code"


rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        Call set_enable_menu(rs.Fields("form_name").Value, _
                rs.Fields("allow_read").Value)
        rs.MoveNext
    Wend
End If
End Sub

Private Sub load_all_menu()
On Error GoTo err_handle

Dim rs As New ADODB.Recordset
Dim i As Integer

rs.Open "select * from m_form order by form_code", CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then rs.MoveFirst

int_count_menu = rs.RecordCount
i = 0
While Not rs.EOF
    i = i + 1
    str_form_title(i) = rs.Fields("form_title").Value
    'Set mnu(i) = Controls(rs.Fields("form_name").Value)
    rs.MoveNext
Wend

Exit Sub
err_handle:
MsgBox "There are some data user/form is invalid!" _
    & vbCrLf & "Please check & update...", vbInformation, headerMSG
End Sub

Private Sub enable_all_menu()
On Error Resume Next

Dim i As Integer
For i = 1 To int_count_menu
    'Call set_enable_menu(mnu(i).name, True)
Next
End Sub

Private Sub disable_all_menu()
On Error Resume Next

Dim i As Integer
For i = 1 To int_count_menu
    'Call set_enable_menu(mnu(i).name, False)
Next
End Sub

Private Sub set_enable_menu(ByVal str1 As String, ByVal blnEnable As Boolean)
'On Error Resume Next
'
'Dim i As Integer
'
'For i = 1 To int_count_menu
'    If mnu(i).name = str1 Then
'        mnu(i).Enabled = blnEnable
'    End If
'Next i
End Sub

Private Sub set_enable_menu_bak(ByRef mnu1 As Menu, ByVal blnEnable As Boolean)
mnu1.Enabled = blnEnable
End Sub

Private Sub timer_terminate_Timer()
End
End Sub

'Private Sub mdi_absensi_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Cancel = -1 'untuk mendisable Alt + F4 </i>
'End Sub

