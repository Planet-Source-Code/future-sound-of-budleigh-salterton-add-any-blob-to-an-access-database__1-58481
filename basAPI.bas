Attribute VB_Name = "basAPI"

Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias _
        "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
        "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
        ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4




Const MAX_PATH = 255
Public Function LaunchAppFile(f As Form, fName As String)
ShellExecute f.hwnd, vbNullString, fName, vbNullString, App.Path, SW_SHOWNORMAL
End Function
Public Function GetTempDir() As String
    Dim sRet As String, lngLen As Long
    
    'create buffer
    sRet = String(MAX_PATH, 0)

    lngLen = GetTempPath(MAX_PATH, sRet)
    If lngLen = 0 Then err.Raise err.LastDllError
    GetTempDir = Left$(sRet, lngLen)
End Function

Public Function AddOfficeBorder(ByVal hwnd As Long)
    
    Dim lngRetVal As Long
    
    'Retrieve the current border style
    lngRetVal = GetWindowLong(hwnd, GWL_EXSTYLE)
    
    'Calculate border style to use
    lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    
    'Apply the changes
    SetWindowLong hwnd, GWL_EXSTYLE, lngRetVal
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Function

