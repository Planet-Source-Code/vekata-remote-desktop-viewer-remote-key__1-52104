VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmImage 
   Caption         =   "Image"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsckImage 
      Left            =   4200
      Tag             =   "0"
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image imgImage 
      Height          =   4815
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Me.BackColor = vbWhite
'Me.KeyPreview = True
wsckImage.Close
wsckImage.Connect FrmMain.txtIP, 53123
End Sub

Private Sub Form_Resize()
Me.BorderStyle = 2
Me.ClipControls = True
imgImage.Width = ScaleWidth
imgImage.Height = ScaleHeight
End Sub

Private Sub imgImage_DblClick()
On Error Resume Next
WindowState = 2
Me.BorderStyle = 0
Me.Width = Screen.Width
Me.Height = Screen.Height
End Sub

Private Sub wsckImage_Close()
Me.BackColor = vbWhite
Print "Connection closed"
End Sub

Private Sub wsckImage_Connect()
Me.BackColor = vbBlack
End Sub

Private Sub wsckImage_DataArrival(ByVal bytesTotal As Long)
Dim str$
If wsckImage.Tag <> 1 Then
    If Dir("Temp.jpg", vbNormal) <> "" Then
        Kill "Temp.jpg"
    End If
    Open "Temp.jpg" For Binary As #1
    wsckImage.Tag = 1
End If
wsckImage.GetData str, vbString, bytesTotal
Put #1, , str
If Right(str, 4) = "$END" Then
    wsckImage.Tag = 0
    Close #1
    imgImage.Picture = LoadPicture("Temp.jpg")
    wsckImage.SendData "ok"
End If
End Sub

Private Sub wsckImage_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbCritical, App.Title
End Sub
