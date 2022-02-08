Attribute VB_Name = "mdlAPI"
Option Explicit
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const PM_REMOVE = &H1

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type msg
    hwnd As Long
    messAge As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMAx As Long, ByVal wRemoveMsg As Long) As Long

