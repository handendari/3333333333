VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_rpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORT"
   ClientHeight    =   3090
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   4680
   Icon            =   "frm_viewer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frm_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub rpt_view _
(ByVal sql_proc As String, ByVal rpt_file As String, ByVal str_param As String)
Dim CrApp As New CRAXDRT.Application
Dim CrRep As New CRAXDRT.Report
Dim AdoRs As New ADODB.Recordset

AdoRs.Open sql_proc, CnG, adOpenDynamic, adLockBatchOptimistic
Set CrRep = CrApp.OpenReport(App.Path & rpt_file)
CrRep.DiscardSavedData
CrRep.Database.Tables(1).SetDataSource AdoRs, 3
CRV.ReportSource = CrRep
CRV.ViewReport

CrRep.ParameterFields.GetItemByName("p_periode").AddCurrentValue str_param
End Sub

Public Sub rpt_view_master _
(ByVal sql_proc As String, ByVal rpt_file As String)
Dim CrApp As New CRAXDRT.Application
Dim CrRep As New CRAXDRT.Report
Dim AdoRs As New ADODB.Recordset

AdoRs.Open sql_proc, CnG, adOpenDynamic, adLockBatchOptimistic
Set CrRep = CrApp.OpenReport(App.Path & rpt_file)
CrRep.DiscardSavedData
CrRep.Database.Tables(1).SetDataSource AdoRs, 3
CRV.ReportSource = CrRep
CRV.ViewReport
End Sub

Private Sub Form_Activate()
Me.WindowState = vbMaximized
End Sub

Private Sub Form_Resize()
CRV.Top = 0
CRV.Left = 0
CRV.Width = Me.Width - 200
CRV.Height = Me.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.WindowState = vbNormal
End Sub
