VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long


Const MAX_PATH = 260
Const BIF_RETURNONLYFSDIRS = 1

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'-----------------------------

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Function Folder(hwnd As Long)
On Error GoTo Out
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim spath As String, udtBI As BrowseInfo
    Dim i As Integer
    
    With udtBI
        .hWndOwner = hwnd
        .lpszTitle = lstrcat("Pilih lokasi foldernya", "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        spath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, spath
        CoTaskMemFree lpIDList
        iNull = InStr(spath, vbNullChar)
        If iNull Then
            spath = Left$(spath, iNull - 1)
        End If
    End If
    Folder = spath
Out:
End Function

Function OpenFlDlg(ByVal HwndOwn As Long, ByVal FilterStr$, ByVal TitleOpen$, ByVal initDir$, Optional MultiSelect As Boolean = False) As String

    Dim ofn As OPENFILENAME
    Dim lpstrz As String * 4480
    Dim tmp$()
    Dim tmp2$
    Dim i&
    
    FilterStr$ = Replace(FilterStr$, "|", Chr$(0))
    ofn.lpstrTitle = TitleOpen
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = HwndOwn
    ofn.hInstance = App.hInstance
    ofn.lpstrFile = lpstrz
    ofn.lpstrFile = String$(4480, 0)
    ofn.nMaxFile = 255
    ofn.lpstrFilter = FilterStr$
    ofn.nFilterIndex = 1
    ofn.lpstrFileTitle = vbNullString
    ofn.nMaxFileTitle = 0
    ofn.lpstrInitialDir = initDir
    ofn.flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or IIf(MultiSelect, OFN_ALLOWMULTISELECT Or OFN_EXPLORER, 0)
    
    If GetOpenFileName(ofn) <> 0 Then
        If Not MultiSelect Then
          OpenFlDlg = Left(ofn.lpstrFile, lstrlen(ofn.lpstrFile))
        Else
          tmp = Split(ofn.lpstrFile, Chr$(0))
          For i = LBound(tmp) To UBound(tmp)
            If Len(tmp(i)) > 0 Then
               tmp2 = tmp2 & tmp(i) & Chr$(0)
            End If
          Next i
          OpenFlDlg = Left$(tmp2, Len(tmp2) - 1)
        End If
    Else
        OpenFlDlg = ""
    End If

End Function