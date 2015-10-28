VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_progess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processing..."
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
End
Attribute VB_Name = "frm_progess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim start_rec, running_rec, end_rec As Integer
Dim bln_activated As Boolean


Public Sub show_progrees()
On Error Resume Next
ProgressBar1.Value = (running_rec / end_rec) * 100
End Sub

Private Sub get_fpdata()
On Error GoTo err_handle

Dim rs As New ADODB.Recordset
Dim a As Integer
a = frm_mst_enroll.get_all_user_info

rs.Open "select enrollnumber from m_enroll where ifnull(fpdata_created,0)=0 and ip_address = '" _
& FG_IP_ADDRESS & "'" _
& " order by enrollnumber asc", CnG, adOpenKeyset, adLockOptimistic

If rs.RecordCount > 0 Then
    end_rec = rs.RecordCount
    rs.MoveFirst
    start_rec = rs.Fields("enrollnumber").Value
    
    While Not rs.EOF
        running_rec = rs.Fields("enrollnumber").Value
        Call frm_mst_enroll.download_enroll(running_rec)
        Call show_progrees
        
        CnG.Execute "update m_enroll set fpdata_created=1 where enrollnumber=" & rs.Fields("enrollnumber").Value _
            & " and ip_address = '" & FG_IP_ADDRESS & "'"
        rs.MoveNext
    Wend
End If

rs.Close
MsgBox end_rec & " enrolls data was succesfully recorded...", vbInformation, headerMSG
Unload Me

Exit Sub

err_handle:
MsgBox "Error downloading Finger data!", vbInformation, headerMSG
Unload Me
End Sub

Private Sub Form_Activate()
If bln_activated = False Then
    bln_activated = True
    Timer1.Enabled = True
End If
End Sub

Private Sub Form_Load()
bln_activated = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call frm_mst_enroll.load_data_enroll
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Call get_fpdata
'MsgBox "OK"
'Unload Me
End Sub
