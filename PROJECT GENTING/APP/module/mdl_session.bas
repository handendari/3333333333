Attribute VB_Name = "mdl_session"
Option Explicit
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
    ByVal idHook As Long, _
    ByVal lpfn As Long, _
    ByVal hmod As Long, _
    ByVal dwThreadId As Long _
) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
    ByVal hHook As Long _
) As Long
Private Declare Function CallNextHookEx Lib "user32" ( _
    ByVal hHook As Long, _
    ByVal ncode As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long
Private Const WH_KEYBOARD = 2
Private Const WH_MOUSE = 7
Private isKBDhooked As Boolean
Private hkbdhook As Long
Private isMousehooked As Boolean
Private hMousehook As Long
Public Sub SetKeyboardHook()
    If isKBDhooked Then Exit Sub
    hkbdhook = SetWindowsHookEx(WH_KEYBOARD, AddressOf WndKeyBoardProc, 0, App.ThreadID)
    isKBDhooked = True
End Sub
Public Sub RemoveKeyboardHook()
    UnhookWindowsHookEx hkbdhook
    isKBDhooked = False
End Sub
Public Function WndKeyBoardProc(ByVal uCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    mdi_absensi.ResetTimer
    WndKeyBoardProc = CallNextHookEx(hkbdhook, uCode, wParam, lParam)
End Function
Public Sub SetMouseHook()
    If isMousehooked Then Exit Sub
    hMousehook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0, App.ThreadID)
    isMousehooked = True
End Sub
Public Sub RemoveMouseHook()
    UnhookWindowsHookEx hMousehook
    isMousehooked = False
End Sub
Public Function MouseProc(ByVal uCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    mdi_absensi.ResetTimer
    MouseProc = CallNextHookEx(hMousehook, uCode, wParam, lParam)
End Function

