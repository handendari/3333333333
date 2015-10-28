VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HDSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const VER_PLATFORM_WIN32S = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const DFP_RECEIVE_DRIVE_DATA = &H7C088
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const CREATE_NEW = 1

Private Enum HDINFO
    HD_MODEL_NUMBER
    HD_SERIAL_NUMBER
    HD_FIRMWARE_REVISION
End Enum

Private Type OSVersionInfo
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type IDEREGS
    bFeaturesReg As Byte
    bSectorCountReg As Byte
    bSectorNumberReg As Byte
    bCylLowReg As Byte
    bCylHighReg As Byte
    bDriveHeadReg As Byte
    bCommandReg As Byte
    bReserved As Byte
End Type

Private Type SENDCMDINPARAMS
    cBufferSize As Long
    irDriveRegs As IDEREGS
    bDriveNumber As Byte
    bReserved(1 To 3) As Byte
    dwReserved(1 To 4) As Long
End Type

Private Type DRIVERSTATUS
    bDriveError As Byte
    bIDEStatus As Byte
    bReserved(1 To 2) As Byte
    dwReserved(1 To 2) As Long
End Type

Private Type SENDCMDOUTPARAMS
    cBufferSize As Long
    DStatus As DRIVERSTATUS
    bBuffer(1 To 512) As Byte
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersionInfo) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, _
    ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetLastError Lib "kernel32" () As Long

Private mvarCurrentDrive As Byte
Private mvarPlatform As String


Public Property Get Copyright() As String
Copyright = "HDSN Vrs. 1.00, (C) Antonio Giuliana, 2001-2003"
End Property

Public Function GetModelNumber() As String
GetModelNumber = CmnGetHDData(HD_MODEL_NUMBER)
End Function

Public Function GetSerialNumber() As String
GetSerialNumber = CmnGetHDData(HD_SERIAL_NUMBER)
End Function

Public Function GetFirmwareRevision() As String
GetFirmwareRevision = CmnGetHDData(HD_FIRMWARE_REVISION)
End Function

Public Property Let CurrentDrive(ByVal vData As Byte)
If vData < 0 Or vData > 3 Then
    Err.Raise 10000, , "Illegal drive number"
End If
mvarCurrentDrive = vData
End Property

Public Property Get CurrentDrive() As Byte
CurrentDrive = mvarCurrentDrive
End Property

Public Property Get Platform() As String
Platform = mvarPlatform
End Property

Private Sub Class_Initialize()
Dim OS As OSVersionInfo
    
OS.dwOSVersionInfoSize = Len(OS)
Call GetVersionEx(OS)
mvarPlatform = "Unk"
Select Case OS.dwPlatformId
    Case Is = VER_PLATFORM_WIN32S
        mvarPlatform = "32S"                ' Win32S
    Case Is = VER_PLATFORM_WIN32_WINDOWS
        If OS.dwMinorVersion = 0 Then
            mvarPlatform = "W95"            ' Win 95
        Else
            mvarPlatform = "W98"            ' Win 98
        End If
    Case Is = VER_PLATFORM_WIN32_NT
        mvarPlatform = "WNT"                ' Win NT/2000
End Select
End Sub

Private Function CmnGetHDData(hdi As HDINFO) As String
Dim bin As SENDCMDINPARAMS
Dim bout As SENDCMDOUTPARAMS
Dim hdh As Long
Dim br As Long
Dim ix As Long
Dim hddfr As Long
Dim hddln As Long
Dim s As String

Select Case hdi
    Case HD_MODEL_NUMBER
        hddfr = 55
        hddln = 40
    Case HD_SERIAL_NUMBER
        hddfr = 21
        hddln = 20
    Case HD_FIRMWARE_REVISION
        hddfr = 47
        hddln = 8
    Case Else
        Err.Raise 10001, "Illegal HD Data type"
End Select

Select Case mvarPlatform
    Case "WNT"
        hdh = CreateFile("\\.\PhysicalDrive" & mvarCurrentDrive, _
            GENERIC_READ + GENERIC_WRITE, FILE_SHARE_READ + FILE_SHARE_WRITE, _
            0, OPEN_EXISTING, 0, 0)
    Case "W95", "W98"
        hdh = CreateFile("\\.\Smartvsd", _
            0, 0, 0, CREATE_NEW, 0, 0)
    Case Else
        Err.Raise 10002, , "Illegal platform (only WNT, W98 or W95)"
End Select
If hdh = 0 Then
    Err.Raise 10003, , "Error on CreateFile"
End If

ZeroMemory bin, Len(bin)
ZeroMemory bout, Len(bout)

With bin
    .bDriveNumber = mvarCurrentDrive
    .cBufferSize = 512
    With .irDriveRegs
        If (mvarCurrentDrive And 1) Then
            .bDriveHeadReg = &HB0
        Else
            .bDriveHeadReg = &HA0
        End If
        .bCommandReg = &HEC
        .bSectorCountReg = 1
        .bSectorNumberReg = 1
    End With
End With

DeviceIoControl hdh, DFP_RECEIVE_DRIVE_DATA, bin, Len(bin), bout, Len(bout), br, 0

s = ""
For ix = hddfr To hddfr + hddln - 1 Step 2
    If bout.bBuffer(ix + 1) = 0 Then Exit For
    s = s & Chr(bout.bBuffer(ix + 1))
    If bout.bBuffer(ix) = 0 Then Exit For
    s = s & Chr(bout.bBuffer(ix))
Next ix

CloseHandle hdh
CmnGetHDData = Trim(s)
End Function
