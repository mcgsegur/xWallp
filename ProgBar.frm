VERSION 5.00
Begin VB.Form ProgBar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "ProgBar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "66%"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "ProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Me.Move 100, 100
  Me.Visible = True
End Sub
Public Function Indica(v As Single, vMax As Single, texto As String)
  Dim porcento As Single
  porcento = (v / vMax)
  Label2.Width = Label1.Width * porcento
  If Len(texto) = 0 Then
    Label2.Caption = CInt(100 * porcento) & "%"
  Else
   Label2.Caption = texto
  End If
End Function
Public Function setTexto(txt As String)
  Label2.Caption = txt
End Function

