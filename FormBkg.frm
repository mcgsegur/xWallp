VERSION 5.00
Begin VB.Form FormBkg 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FormBkg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   316
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FormBkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim largTelaTwips As Long, altTelaTwips As Long
Private Sub Form_Load()
  Dim nRetry As Long
  largTelaTwips = screenWidthTwips
  altTelaTwips = screenHeightTwips
  Me.Move 0, 0, largTelaTwips, altTelaTwips
  For nRetry = 0 To 3
    If putformBehindIconesDesktop(Me) Then Exit For
    Pausa 1#
  Next nRetry
  If nRetry = 4 Then End
End Sub
Public Sub CarregaImagemFundo(imagem As StdPicture)
  FormBkg.PaintPicture imagem, 0, 0, largTelaTwips, altTelaTwips, 0, 0, ScaleX(imagem.Width, vbHimetric, vbTwips), ScaleY(imagem.Height, vbHimetric, vbTwips)
  FormBkg.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  endPutFormBehindDesktopIcons
End Sub
