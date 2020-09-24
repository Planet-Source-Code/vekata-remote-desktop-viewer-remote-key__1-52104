VERSION 5.00
Begin VB.Form fSaveJPG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JPEG Compression Settings"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
End
Attribute VB_Name = "fSaveJPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private m_Image     As cImage
Private m_Jpeg      As cJpeg
Private m_FileName  As String


Public Sub SaveImage(TheImage As cImage, FileName As String)
    Set m_Image = TheImage 'Call this before the form loads to initialize it
    m_FileName = FileName
End Sub

Private Sub Form_Load()
    Set m_Jpeg = New cJpeg
    m_Jpeg.Quality = 75
    
    m_Jpeg.SampleHDC m_Image.hDC, m_Image.Width, m_Image.Height

       'Delete file if it exists
        RidFile m_FileName

       'Save the JPG file
        m_Jpeg.SaveFile m_FileName

    Set m_Image = Nothing
    Set m_Jpeg = Nothing
    Unload Me
End Sub
