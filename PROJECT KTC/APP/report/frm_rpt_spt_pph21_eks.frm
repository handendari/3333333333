VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_rpt_spt_pph21_eks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SPT PPH PASAL 21 KARYAWAN YANG TELAH BERHENTI"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin prj_fej_jkt.LynxGrid LynxGrid2 
      Height          =   2895
      Left            =   1950
      TabIndex        =   2
      Top             =   1980
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5106
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
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      AllowColumnResizing=   -1  'True
      ColumnSort      =   -1  'True
   End
   Begin VB.Frame Frame2 
      Caption         =   "Yearly"
      Height          =   3555
      Left            =   150
      TabIndex        =   4
      Top             =   180
      Width           =   8025
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         Top             =   720
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   118226947
         CurrentDate     =   40904
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frm_rpt_spt_pph21_eks.frx":0000
         Left            =   1800
         List            =   "frm_rpt_spt_pph21_eks.frx":000A
         TabIndex        =   9
         Top             =   1110
         Width           =   1605
      End
      Begin VB.TextBox txtnmkar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         DragMode        =   1  'Automatic
         Height          =   285
         Left            =   3300
         TabIndex        =   7
         Top             =   1500
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox txtkdkar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1500
         Visible         =   0   'False
         Width           =   1125
      End
      Begin prj_fej_jkt.vbButton vbButton2 
         Height          =   285
         Left            =   2940
         TabIndex        =   5
         Top             =   1500
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "..."
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
         MICON           =   "frm_rpt_spt_pph21_eks.frx":001D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Year :"
         Height          =   210
         Left            =   870
         TabIndex        =   11
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Employee :"
         Height          =   210
         Left            =   870
         TabIndex        =   8
         Top             =   1140
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Height          =   0
      Left            =   150
      TabIndex        =   1
      Top             =   4110
      Width           =   8055
   End
   Begin prj_fej_jkt.vbButton vbButton5 
      Height          =   585
      Left            =   6390
      TabIndex        =   0
      Top             =   3990
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1032
      BTYPE           =   14
      TX              =   "EXIT"
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
      MICON           =   "frm_rpt_spt_pph21_eks.frx":0039
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_fej_jkt.vbButton vbButton1 
      Height          =   585
      Left            =   4275
      TabIndex        =   3
      Top             =   4005
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1032
      BTYPE           =   14
      TX              =   "Print"
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
      MICON           =   "frm_rpt_spt_pph21_eks.frx":0055
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
Attribute VB_Name = "frm_rpt_spt_pph21_eks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim strsql As String
Public editTrans As Boolean

Private Sub createKar()
   With LynxGrid2
      .AddColumn "NIK", 700, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "Department", , , , , , , , , False
      .AddColumn "Tittle", 2000, , , , , , , , , True
      .AddColumn "Location", , , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear
        strsql = "select employee_code,employee_name,department,title," _
                    & "location " _
                & "from h_salary_eks " _
                & "WHERE (employee_code LIKE '%" & txtkdkar.Text & "%' " _
                & "OR employee_name LIKE '%" & txtkdkar.Text & "%') ORDER BY employee_name"
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs2.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid2.AddItem rs2!employee_code & vbTab & rs2!employee_name & vbTab & _
                rs2!department & vbTab & rs2!TITLE & vbTab & rs2!Location
                rs2.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs2.RecordCount = 1 Then
                rs2.MoveFirst
                txtkdkar.Text = rs2!employee_code
                txtnmkar.Text = rs2!employee_name
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else
            
        End If
        rs2.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txtkdkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txtnmkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 1 Then
    txtkdkar.Visible = True
    vbButton2.Visible = True
    txtnmkar.Visible = True
Else
    txtkdkar.Visible = False
    txtkdkar.Text = ""
    vbButton2.Visible = False
    txtnmkar.Visible = False
    txtnmkar.Text = ""
End If
End Sub

Private Sub Form_Load()

createKar
DTPicker1.Value = Now
Combo1.ListIndex = 0
    
End Sub

Private Sub LynxGrid2_DblClick()
    isiGridKar (2)
End Sub

Private Sub LynxGrid2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        LynxGrid2.Visible = False
    End If
    If KeyAscii = 13 Then
        isiGridKar (2)
    End If
End Sub

Private Sub LynxGrid2_LostFocus()
    LynxGrid2.Visible = False
End Sub

Private Sub txtkdkar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub vbButton1_Click()
Dim str_sql, str_param_periode, str_file As String
Dim a As New frm_rpt
Dim d1, d2 As String

str_file = "\report\rpt_spt_pph21_eks.rpt"
str_employee_code = txtkdkar.Text


d1 = Format(DTPicker1.Value, "yyyy")

str_sql = "CALL spr_pph21_eks('" & str_employee_code & "',  '" & d1 & "', " & _
        "" & Combo1.ListIndex & ",'" & LOGIN_LEVEL & "');"
str_param_periode = "YEARLY : (" & Format(DTPicker1.Value, "yyyy") & ")"

Call a.Show
a.Caption = "SPT PPH PASAL 21 REPORT KARYAWAT YANG TELAH BERHENTI"
Call a.rpt_view_spt_pph21(str_sql, str_file, str_param_periode)

End Sub

Private Sub vbButton2_Click()
    isiGridKar (1)
End Sub

Private Sub vbButton5_Click()
    Unload Me
End Sub
