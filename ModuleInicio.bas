Attribute VB_Name = "ModuleInicio"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Public Const SPIF_UPDATEINIFILE = 1
Public Const SPIF_SENDCHANGE = 2

Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_GETDESKWALLPAPER = &H73



Public Const CADINDSTARTUP = "--@startup"
Public Const CADINDNORUN = "--@norun"

'Public pathDebug As String
Public appPath As String
Public bInsideIDE As Boolean

'Public Sub geralogDebug(msg As String)
'  Dim ifile As Integer
'  ifile = FreeFile
'  Open pathDebug For Append As #ifile
'  Print #ifile, msg
'  Close #ifile
'
'End Sub


Sub Main()
  Dim sArgs() As String
  Dim cnt As Long
  If App.PrevInstance Then Exit Sub
  
  On Error Resume Next
  Debug.Print 3 / 0
  bInsideIDE = (Err.Number <> 0)
  On Error GoTo 0
  appPath = App.path
  If Right$(appPath, 1) <> "\" Then appPath = appPath & "\"


'  Muda_prioridade REALTIME_PRIORITY_CLASS, THREAD_PRIORITY_TIME_CRITICAL
'  Muda_prioridade HIGH_PRIORITY_CLASS, THREAD_PRIORITY_HIGHEST
  Muda_prioridade NORMAL_PRIORITY_CLASS, THREAD_PRIORITY_IDLE

  Debug.Print "Prioridades: Processo= " & Le_prioridade_processo & "   Thread=" & Le_prioridade_thread
  If Command$ <> "" Then
    sArgs = Split(Command$, " ")
    If LCase$(sArgs(0)) = CADINDNORUN Then End

    If LCase$(sArgs(0)) = CADINDSTARTUP Then
      ' se for ativado quando 'logar'(startup da sessão), espera PROGMAN 'se acertar'
      Do
        Sleep 500
        DoEvents
      Loop While GetHandleWProgMan() = 0
      ' espero mais 3 segundos....
      For cnt = 1 To 30 ' tenta garantir que programa vai 'responder' às mensagens do windows
        Sleep 100
        DoEvents
      Next cnt
    End If
  End If
  Load Form1
End Sub
