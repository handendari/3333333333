VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_rpt2 
   Caption         =   "REPORT"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   Icon            =   "frm_viewer2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   6780
   WindowState     =   2  'Maximized
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
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
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
Attribute VB_Name = "frm_rpt2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CrApp As New CRAXDRT.Application
Dim CrRep As New CRAXDRT.Report
Dim AdoRs As New ADODB.Recordset

Public Sub viewReport _
(ByRef sql_proc, sql_proc2, sql_proc3, rpt_file, pnmUsaha, pAlamat, pKota, ptlp, puser, str_param1, str_param2, _
    str_param3 As String)
'Dim CrApp As New CRAXDRT.Application
'Dim CrRep As New CRAXDRT.Report

Dim CrDatabase As CRAXDRT.Database
Dim CrDatabaseTables As CRAXDRT.DatabaseTables
Dim CrDatabaseTable As CRAXDRT.DatabaseTable
Dim CrSections As CRAXDRT.Sections
Dim CrSection As CRAXDRT.Section
Dim CrReportObjs As CRAXDRT.ReportObjects
Dim CrSubreportObj As CRAXDRT.SubreportObject
Dim CrSubreport As CRAXDRT.Report


Dim AdoRs As New ADODB.Recordset
Dim AdoRs2 As New ADODB.Recordset
Dim Adors3 As New ADODB.Recordset

    Me.MousePointer = vbHourglass
    
    AdoRs.Open sql_proc, CnG, adOpenForwardOnly, adLockReadOnly
    AdoRs2.Open sql_proc2, CnG, adOpenForwardOnly, adLockReadOnly
    Adors3.Open sql_proc3, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set CrRep = CrApp.OpenReport(App.Path & rpt_file)
    CrRep.DiscardSavedData
    CrRep.Database.Tables(1).SetDataSource AdoRs, 3
    
'    CrRep.ParameterFields.GetItemByName("p_nmUsaha").AddCurrentValue pnmUsaha
'    CrRep.ParameterFields.GetItemByName("p_alamat").AddCurrentValue pAlamat
'    CrRep.ParameterFields.GetItemByName("p_kota").AddCurrentValue pKota
'    CrRep.ParameterFields.GetItemByName("p_telepon").AddCurrentValue ptlp
'    CrRep.ParameterFields.GetItemByName("p_user").AddCurrentValue puser
'
    CrRep.ParameterFields.GetItemByName("p_param1").AddCurrentValue str_param1
'    CrRep.ParameterFields.GetItemByName("p_param2").AddCurrentValue str_param2
'    CrRep.ParameterFields.GetItemByName("p_param3").AddCurrentValue str_param3

Dim x As Integer
Dim Y As Integer
Dim tabel As String

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
                tabel = CrDatabaseTables.item(1).name
                If tabel = "ttx_slip_gaji_detail2A_ttx" Then
                    CrDatabaseTable.SetDataSource Adors3, 3
                Else
                    CrDatabaseTable.SetDataSource AdoRs2, 3
                End If
                
                
                'MsgBox "Table : " & CrDatabaseTable.Name & " in subreport " _
                    & UCase(CrSubreportObj.SubreportName) & " in location " _
                    & CrDatabaseTable.Location
                
            End If
        Next
           
    Next
    
    crv.ReportSource = CrRep
    crv.viewReport
    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
    crv.Top = 0
    crv.Left = 0
    crv.Width = Me.Width - 200
    crv.Height = Me.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_rpt = Nothing
End Sub

Private Sub crv_PrintButtonClicked(UseDefault As Boolean)
    UseDefault = False
    CrRep.PrinterSetup Me.hWnd
    CrRep.PrintOut True
    'Set CrRep = Nothing
    'crv.Visible = False
End Sub
