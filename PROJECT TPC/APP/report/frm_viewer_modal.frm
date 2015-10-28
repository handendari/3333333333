VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_rpt_modal 
   Caption         =   "REPORT"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frm_viewer_modal.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
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
Attribute VB_Name = "frm_rpt_modal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CrApp As New CRAXDRT.Application
Dim CrRep As New CRAXDRT.Report
Dim AdoRs As New ADODB.Recordset

Public Sub rpt_view _
(ByVal sql_proc As String, ByVal rpt_file As String, ByVal str_param As String)
'Dim CrApp As New CRAXDRT.Application
'Dim CrRep As New CRAXDRT.Report
'Dim AdoRs As New ADODB.Recordset

AdoRs.Open sql_proc, CnG, adOpenDynamic, adLockBatchOptimistic
Set CrRep = CrApp.OpenReport(App.Path & rpt_file)
CrRep.DiscardSavedData
CrRep.Database.Tables(1).SetDataSource AdoRs, 3
crv.ReportSource = CrRep
crv.ViewReport

CrRep.ParameterFields.GetItemByName("p_periode").AddCurrentValue str_param
End Sub

Public Sub rpt_print _
(ByVal sql_proc As String, ByVal rpt_file As String, ByVal str_param As String)
'Dim CrApp As New CRAXDRT.Application
'Dim CrRep As New CRAXDRT.Report
'Dim AdoRs As New ADODB.Recordset

AdoRs.Open sql_proc, CnG, adOpenDynamic, adLockBatchOptimistic
Set CrRep = CrApp.OpenReport(App.Path & rpt_file)
CrRep.DiscardSavedData
CrRep.Database.Tables(1).SetDataSource AdoRs, 3
crv.ReportSource = CrRep
'CRV.ViewReport
'CRV.Visible = False
CrRep.PrintOut True
'crv_PrintButtonClicked (False)

CrRep.ParameterFields.GetItemByName("p_periode").AddCurrentValue str_param
End Sub

Public Sub rpt_view_spt_pph21 _
(ByVal sql_proc As String, ByVal rpt_file As String, ByVal str_param As String)
'Dim CrApp As New CRAXDRT.Application
'Dim CrRep As New CRAXDRT.Report
'Dim AdoRs As New ADODB.Recordset

AdoRs.Open sql_proc, CnG, adOpenDynamic, adLockBatchOptimistic
Set CrRep = CrApp.OpenReport(App.Path & rpt_file)
CrRep.DiscardSavedData
CrRep.Database.Tables(1).SetDataSource AdoRs, 3
crv.ReportSource = CrRep
crv.ViewReport

'CrRep.ParameterFields.GetItemByName("p_periode").AddCurrentValue str_param
End Sub

Public Sub rpt_view_master _
(ByVal sql_proc As String, ByVal rpt_file As String)
'Dim CrApp As New CRAXDRT.Application
'Dim CrRep As New CRAXDRT.Report
'Dim AdoRs As New ADODB.Recordset

AdoRs.Open sql_proc, CnG, adOpenDynamic, adLockBatchOptimistic
Set CrRep = CrApp.OpenReport(App.Path & rpt_file)
CrRep.DiscardSavedData
CrRep.Database.Tables(1).SetDataSource AdoRs, 3
crv.ReportSource = CrRep
crv.ViewReport
End Sub

Private Sub Form_Activate()
Me.WindowState = vbMaximized
End Sub

Private Sub Form_Resize()
crv.Top = 0
crv.Left = 0
crv.Width = Me.Width - 200
crv.Height = Me.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.WindowState = vbNormal
End Sub

'===
'===
Public Sub rpt_auto_pdf(ByVal sql_proc As String, ByVal rpt_file As String, ByVal str_param As String)
'Dim CrApp As New CRAXDRT.Application
'Dim CrRep As New CRAXDRT.Report
'Dim AdoRs As New ADODB.Recordset

AdoRs.Open sql_proc, CnG, adOpenDynamic, adLockBatchOptimistic
Set CrRep = CrApp.OpenReport(App.Path & rpt_file)
CrRep.DiscardSavedData
CrRep.Database.Tables(1).SetDataSource AdoRs, 3

CrRep.ParameterFields.GetItemByName("p_periode").AddCurrentValue str_param

'---
CrRep.ExportOptions.DestinationType = crEDTDiskFile
CrRep.ExportOptions.FormatType = crEFTPortableDocFormat
CrRep.ExportOptions.DiskFileName = App.Path & "\report\test.pdf"
CrRep.Export False
End Sub


Public Sub rpt_view2 _
(ByVal sql1 As String, ByVal sql2 As String, ByVal Data2 As String, _
ByVal rpt_file As String, ByVal str_param As String)

'Dim CrApp As New CRAXDRT.Application
'Dim CrRep As CRAXDRT.Report

Dim CrDatabase As CRAXDRT.Database
Dim CrDatabaseTables As CRAXDRT.DatabaseTables
Dim CrDatabaseTable As CRAXDRT.DatabaseTable
Dim CrSections As CRAXDRT.Sections
Dim CrSection As CRAXDRT.Section
Dim CrReportObjs As CRAXDRT.ReportObjects
Dim CrSubreportObj As CRAXDRT.SubreportObject
Dim CrSubreport As CRAXDRT.Report

Dim AdoRs1 As New ADODB.Recordset
Dim AdoRs2 As New ADODB.Recordset

' ------
Dim x As Integer
Dim Y As Integer

AdoRs1.Open sql1, CnG, adOpenDynamic, adLockBatchOptimistic
AdoRs2.Open sql2, CnG, adOpenDynamic, adLockBatchOptimistic

Set CrRep = CrApp.OpenReport(App.Path & rpt_file)

CrRep.DiscardSavedData
CrRep.Database.Tables(1).SetDataSource AdoRs1, 3
'--

Set CrSections = CrRep.Sections
For x = 1 To CrSections.Count
    Set CrSection = CrSections.item(x)
    Set CrReportObjs = CrSection.ReportObjects
    For Y = 1 To CrReportObjs.Count
        If CrReportObjs.item(Y).Kind = crSubreportObject Then
        
            Set CrSubreportObj = CrReportObjs.item(Y)
            Set CrSubreport = CrSubreportObj.OpenSubreport
            Set CrDatabase = CrSubreport.Database

            Set CrDatabaseTables = CrDatabase.Tables
            Set CrDatabaseTable = CrDatabaseTables.item(1)
            
            If UCase(CrDatabaseTable.name) = UCase(Data2) Then
                CrDatabaseTable.SetDataSource AdoRs2, 3
'            ElseIf UCase(CrDatabaseTable.Name) = UCase(data2) Then
'                CrDatabaseTable.SetDataSource AdoRs2, 3
            End If
            
            
            'MsgBox "Table : " & CrDatabaseTable.Name & " in subreport " _
                & UCase(CrSubreportObj.SubreportName) & " in location " _
                & CrDatabaseTable.Location
            
        End If
    Next
       
Next


If rpt_file = "\report\rpt_03.rpt" Or rpt_file = "\report\rpt_04.rpt" Then
    CrRep.ParameterFields.GetItemByName("p_periode").AddCurrentValue str_param
End If


crv.ReportSource = CrRep
crv.ViewReport

'CrRep.ParameterFields.GetItemByName("p_periode").AddCurrentValue str_param
End Sub

Public Sub rpt_view3 _
(ByVal sql1 As String, ByVal sql2 As String, ByVal sql3 As String, ByVal str1 As String, ByVal str2 As String, _
ByVal rpt_file As String, ByVal str_param As String)

'Dim CrApp As New CRAXDRT.Application
'Dim CrRep As CRAXDRT.Report

Dim CrDatabase As CRAXDRT.Database
Dim CrDatabaseTables As CRAXDRT.DatabaseTables
Dim CrDatabaseTable As CRAXDRT.DatabaseTable
Dim CrSections As CRAXDRT.Sections
Dim CrSection As CRAXDRT.Section
Dim CrReportObjs As CRAXDRT.ReportObjects
Dim CrSubreportObj As CRAXDRT.SubreportObject
Dim CrSubreport As CRAXDRT.Report

Dim AdoRs1 As New ADODB.Recordset
Dim AdoRs2 As New ADODB.Recordset
Dim AdoRs3 As New ADODB.Recordset

' ------
Dim x As Integer
Dim Y As Integer

AdoRs1.Open sql1, CnG, adOpenDynamic, adLockBatchOptimistic
AdoRs2.Open sql2, CnG, adOpenDynamic, adLockBatchOptimistic
AdoRs3.Open sql3, CnG, adOpenDynamic, adLockBatchOptimistic

Set CrRep = CrApp.OpenReport(App.Path & rpt_file)

CrRep.DiscardSavedData
CrRep.Database.Tables(1).SetDataSource AdoRs1, 3
'--

Set CrSections = CrRep.Sections
For x = 1 To CrSections.Count
    Set CrSection = CrSections.item(x)
    Set CrReportObjs = CrSection.ReportObjects
    For Y = 1 To CrReportObjs.Count
        If CrReportObjs.item(Y).Kind = crSubreportObject Then
        
            Set CrSubreportObj = CrReportObjs.item(Y)
            Set CrSubreport = CrSubreportObj.OpenSubreport
            Set CrDatabase = CrSubreport.Database

            Set CrDatabaseTables = CrDatabase.Tables
            Set CrDatabaseTable = CrDatabaseTables.item(1)
            
            If UCase(CrDatabaseTable.name) = UCase(str1) Then
                CrDatabaseTable.SetDataSource AdoRs2, 3
            ElseIf UCase(CrDatabaseTable.name) = UCase(str2) Then
                CrDatabaseTable.SetDataSource AdoRs3, 3
            End If
            
            
            'MsgBox "Table : " & CrDatabaseTable.Name & " in subreport " _
                & UCase(CrSubreportObj.SubreportName) & " in location " _
                & CrDatabaseTable.Location
            
        End If
    Next
       
Next


If rpt_file = "\report\rpt_03.rpt" Or rpt_file = "\report\rpt_04.rpt" Then
    CrRep.ParameterFields.GetItemByName("p_periode").AddCurrentValue str_param
End If

crv.ReportSource = CrRep
crv.ViewReport
'CrRep.ParameterFields.GetItemByName("p_periode").AddCurrentValue str_param
End Sub

Public Sub crv_PrintButtonClicked(UseDefault As Boolean)
    UseDefault = False
    CrRep.PrinterSetup Me.hwnd
    CrRep.PrintOut True
    'Set CrRep = Nothing
    'crv.Visible = False
End Sub
