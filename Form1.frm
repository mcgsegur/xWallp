VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerTentaDescarregar 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   3960
      Top             =   1680
   End
   Begin VB.Timer TimerTampering 
      Interval        =   1000
      Left            =   1080
      Top             =   720
   End
   Begin VB.Timer TimerIntervalo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   2640
      Top             =   1560
   End
   Begin VB.Menu menuCmds 
      Caption         =   "&Menu"
      Begin VB.Menu mNomeLocal 
         Caption         =   "********"
         Checked         =   -1  'True
      End
      Begin VB.Menu msep1 
         Caption         =   "-"
      End
      Begin VB.Menu mInterval 
         Caption         =   "change picture &Interval"
      End
      Begin VB.Menu msep2 
         Caption         =   "-"
      End
      Begin VB.Menu mOpenPicFolder 
         Caption         =   "&Open pictures folder in explorer"
      End
      Begin VB.Menu mDownloadWpapers 
         Caption         =   "try to &Download more pictures"
      End
      Begin VB.Menu msep3 
         Caption         =   "-"
      End
      Begin VB.Menu mRunStartup 
         Caption         =   "run at windows &Startup"
      End
      Begin VB.Menu msep4 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "&About the program"
      End
      Begin VB.Menu mKeywords 
         Caption         =   "about &Current picture(google)"
      End
      Begin VB.Menu msep5 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const TOESCUROBASE = 80
Const DELTABRIGHTNESS = 5
Const MENORBRIGHTNESS = -48
Const TEMPODESCARGARAPIDA As Single = 1.5 ' tempo em segundos durante o qual tenta descarregar novos WallPapers
Const TEMPODESCARGALONGA As Single = 20#
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, _
     ByVal yPoint As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function Shell_NotifyIcon Lib "Shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205


Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private TrayIcon As NOTIFYICONDATA

Dim lastMousePos As POINTAPI
Dim bescurecendo As Boolean
Dim Brightness As Integer
Dim toEscuro As Long
Dim imagemBase As StdPicture
Dim bColocouIconNaTray As Boolean
Dim pathFolderImg As String
Dim idxArq As Long

Dim intervalo As Long
Dim intervCnt As Long
Dim btestouTampering As Boolean
Dim bEncerrandoPrograma As Boolean
Dim linhasCfg() As String ' copyright/pathFolderImg/Intervalo/livre_futuro/livre_futuro/livre_futuro/livre_futuro/livre_futuro
Dim intervaloTrocas As Long
Dim appDataPath As String
Dim pathCfg As String
Dim divTimerTentaDescarregar As Long
Dim VetIn() As tipoRGBA, VetOut() As tipoRGBA, pWidth As Long, pHeight As Long
Dim bi32BitInfo As BITMAPINFO
Dim nroWp As Long
Dim listaJPegsAvailable() As String
Dim iproxImg As Long
Dim pathArqIdxWallp As String
Dim largTela As Long, altTela As Long
Dim localDaImagemStr As String
Private Sub acertaNroImgsNoMenu()
   mOpenPicFolder.Caption = "&Open pictures folder in explorer (" & nroWp & " available)"
End Sub
Private Sub setIntervaloTroca(v As Long)
  intervalo = v
  mInterval.Caption = "change picture &Interval(" & intervalo & "s)"
  intervCnt = intervalo
  With TimerIntervalo
    .Enabled = False
    .Enabled = (intervalo <> 0)
  End With
  linhasCfg(1) = intervalo
End Sub
Private Function pegaListaJPegsAvailable() As Boolean
  Dim iLivre As Integer, linha As String
  iLivre = FreeFile
  On Error GoTo trataerro
  Open pathArqIdxWallp For Input As #iLivre
  nroWp = 0
  ReDim listaJPegsAvailable(0 To 1000)
  While Not EOF(iLivre)
    Line Input #iLivre, linha
    If Left$(linha, 6) = "!MCGS!" Then
      If Not EOF(iLivre) Then
        nroWp = 0
        Erase listaJPegsAvailable
        Close #iLivre
        Exit Function
      End If
    Else
      If nroWp > UBound(listaJPegsAvailable) Then ReDim Preserve listaJPegsAvailable(0 To nroWp + 300)
      listaJPegsAvailable(nroWp) = Mid$(linha, 45)
      nroWp = nroWp + 1
    End If
  Wend
  ReDim Preserve listaJPegsAvailable(0 To nroWp - 1)
  




'  Dim objFile As Shell32.FolderItem
'  Dim objFolder As Shell32.Folder
'  Dim objshell As New Shell32.Shell
'   On Error GoTo trataerro
'   Set objFolder = objshell.NameSpace(pathFolder)
'   ReDim listaJPegsAvailable(0 To objFolder.Items.Count - 1)
'   nroWp = 0
'   For Each objFile In objFolder.Items
'     If objFile.Type = "JPG File" Then
'       listaJPegsAvailable(nroWp) = objFile.name
'       nroWp = nroWp + 1
'     End If
'   Next objFile
'   ReDim Preserve listaJPegsAvailable(0 To nroWp - 1)
   acertaNroImgsNoMenu
   pegaListaJPegsAvailable = True
   Close #iLivre
   Exit Function
trataerro:
    Close #iLivre
End Function
'Private Function pegaProxArqImg(ByVal bAtualizaImagem As Boolean) As Boolean
'  Dim i As Long, n As Long
'  Dim iBusca As Long, localDaImagemStr As String
''  Dim objshell As New Shell32.Shell
''  Dim objFolder As Shell32.Folder
''  Dim objFile As Shell32.FolderItem
'  Dim larg As Long, alt As Long
'  Dim patharqImg As String
'  Dim nomeArqImg As String
'  Dim idxArq As Long
'  Dim largTela As Long
'   intervCnt = intervalo
'   largTela = screenWidthPixels
'
'
'   n = 0
'   On Error Resume Next
'   For i = 1 To nroWp
'     nomeArqImg = listaJPegsAvailable(iproxImg)
'     Set imagemBase = LoadPicture(pathFolderImg & "\" & nomeArqImg)
'     If Err.Number = 0 Then
'       larg = ScaleX(imagemBase.Width, vbHimetric, vbPixels)
'       alt = ScaleY(imagemBase.Height, vbHimetric, vbPixels)
'       If (larg > alt) And (larg > largTela \ 3) Then
'         ' resgata informações sobre o arquivo ( pelo nome)
'         iBusca = InStrRev(nomeArqImg, "_" & larg & "x" & alt & "_") ' "assinatura"
'         localDaImagemStr = nomeArqImg ' valor default
'         If iBusca <> 0 Then
'           localDaImagemStr = Left$(nomeArqImg, iBusca - 1)
'         End If
'         TrayIcon.szTip = "XWallp:[" & localDaImagemStr & "]" & vbNullChar
'         If bColocouIconNaTray Then
'           Call Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
'         Else
'           If Shell_NotifyIcon(NIM_ADD, TrayIcon) <> 0 Then bColocouIconNaTray = True
'         End If
'         mOpenPicFolder.Caption = "&Open pictures folder in explorer (" & nroWp & " available)"
'         mNomeLocal.Caption = ">> Location:[" & localDaImagemStr & "]"
'
'         pWidth = larg
'         pHeight = alt
'         ReDim VetIn(pWidth * pHeight - 1)
'         ReDim VetOut(pWidth * pHeight - 1)
'         With bi32BitInfo.bmiHeader
'           .biBitCount = 32
'           .biPlanes = 1
'           .biSize = Len(bi32BitInfo.bmiHeader)
'           .biWidth = pWidth
'           .biHeight = pHeight
'           .biSizeImage = 4 * pWidth * pHeight
'         End With
'         Pic.Picture = imagemBase
'         GetDIBits Pic.hdc, Pic.Picture.Handle, 0, pHeight, VetIn(0), bi32BitInfo, 0
'         Pic = LoadPicture()
'         If bAtualizaImagem Then Brightness = -DELTABRIGHTNESS       ' para forçar redesenho....
'         iproxImg = (iproxImg + 1) Mod nroWp ' preparo pegar próximo
'         pegaProxArqImg = True
'         If TimerIntervalo.Enabled Then ' rearma timer
'           TimerIntervalo.Enabled = False
'           TimerIntervalo.Enabled = True
'         End If
'
'         Exit Function
'       End If
'     End If
'   Next i
'   iproxImg = 0
'   ' peguei TODOS e deram erro.....>>>> Falha
'End Function
Private Function pegaProxArqImg(ByVal bAtualizaImagem As Boolean) As Boolean
  Dim i As Long, n As Long
  Dim iBusca As Long
'  Dim objshell As New Shell32.Shell
'  Dim objFolder As Shell32.Folder
'  Dim objFile As Shell32.FolderItem
  Dim largImg As Long, altImg As Long
'  Dim patharqImg As String
  Dim nomeArqImg As String
'  Dim idxArq As Long
  Dim largTela As Long, altTela As Long
   intervCnt = intervalo
   largTela = screenWidthPixels
   altTela = screenHeightPixels
   n = 0
   On Error Resume Next
   For i = 1 To nroWp
     nomeArqImg = listaJPegsAvailable(iproxImg)
     
'     nomeArqImg = "Godrevy__England_1920x1080_F436BBBCF4F6.jpg"
     
     Set imagemBase = LoadPicture(pathFolderImg & "\" & nomeArqImg)
     If Err.Number = 0 Then
       largImg = ScaleX(imagemBase.Width, vbHimetric, vbPixels)
       altImg = ScaleY(imagemBase.Height, vbHimetric, vbPixels)
       If (largImg > altTela) And (largImg > largTela \ 3) Then
         ' resgata informações sobre o arquivo ( pelo nome)
         iBusca = InStrRev(nomeArqImg, "_" & largImg & "x" & altImg & "_") ' "assinatura"
         localDaImagemStr = nomeArqImg ' valor default
         If iBusca <> 0 Then
           localDaImagemStr = Left$(nomeArqImg, iBusca - 1)
         End If
         TrayIcon.szTip = "XWallp:[" & localDaImagemStr & "]" & vbNullChar
         If bColocouIconNaTray Then
           Call Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
         Else
           If Shell_NotifyIcon(NIM_ADD, TrayIcon) <> 0 Then bColocouIconNaTray = True
         End If
         mNomeLocal.Caption = ">> Location:[" & localDaImagemStr & "]"
         If largImg <> largTela Or altImg <> altTela Then ' tenho de redimensionar imagem
           FormBkg.PaintPicture imagemBase, 0, 0, largTela, altTela
           Set imagemBase = FormBkg.image
           largImg = largTela
           altImg = altTela
         End If
         pWidth = largImg
         pHeight = altImg
         
         
         ReDim VetIn(pWidth * pHeight - 1)
         ReDim VetOut(pWidth * pHeight - 1)
         With bi32BitInfo.bmiHeader
           .biBitCount = 32
           .biPlanes = 1
           .biSize = Len(bi32BitInfo.bmiHeader)
           .biWidth = pWidth
           .biHeight = pHeight
           .biSizeImage = 4 * pWidth * pHeight
         End With
         FormBkg.Picture = imagemBase
         GetDIBits FormBkg.hdc, FormBkg.Picture.Handle, 0, pHeight, VetIn(0), bi32BitInfo, 0
         If bAtualizaImagem Then Brightness = -DELTABRIGHTNESS       ' para forçar redesenho....
         iproxImg = (iproxImg + 1) Mod nroWp ' preparo pegar próximo
         pegaProxArqImg = True
         If TimerIntervalo.Enabled Then ' rearma timer
           TimerIntervalo.Enabled = False
           TimerIntervalo.Enabled = True
         End If
         changeCurrentSystemWallPaper pathFolderImg & "\" & nomeArqImg
         Exit Function
       End If
     End If
   Next i
   iproxImg = 0
   ' peguei TODOS e deram erro.....>>>> Falha
End Function




Private Function normaliza(titulo As String) As String
  Dim i As Long, c As String
  normaliza = titulo
'  Debug.Print titulo
  For i = 1 To Len(titulo)
    c = LCase(Mid$(titulo, i, 1))
    If InStr(1, "abcdefghijklmnopqrstuvwxyz0123456789çñãõáéíóúü", c) = 0 Then
      Mid$(normaliza, i, 1) = "_"
    End If
  Next i
End Function

Private Function descarregaWallpapers(duracao As Single, bMostraProgress As Boolean) As Long
 Dim DescrImg As tipoDescrImg
 Dim linhas() As String, tLim As Single
 Dim nomeArqImg As String
' Dim stdPic As StdPicture
 Dim listaSHA256() As String
 Dim bArquivoJaExiste As Boolean
 Dim fullPathArqImg As String
 
' Dim objGDIPlus As New ProcessaImgGDIPlus
 Dim i As Long
 Dim ti As Single
' Dim bSucesso As Boolean
 Dim imgString As String
 Dim iLivre As Integer
 menuCmds.Enabled = False
 descarregaWallpapers = 0
 On Error GoTo trataerro
 
 If leCfgHash(pathArqIdxWallp, linhas, "chvIndice") Then
   nroWp = UBound(linhas) + 1
   ReDim listaSHA256(0 To nroWp + 999)
   For i = 0 To nroWp - 1
     listaSHA256(i) = Left$(linhas(i), 44)
'     nomeArqImg = Mid$(linhas(i), 45)
'     If Dir$(pathFolderImg & "\" & nomeArqImg) = "" Then
'       MsgBox "Não Achei [" & nomeArqImg & "]!"
'     End If
   Next i
   ReDim Preserve linhas(0 To nroWp + 999)
   If GetTipoConexaoWeb = TCI_NOTAVAILABLE Then
     descarregaWallpapers = nroWp
     menuCmds.Enabled = True
     Exit Function
   End If
   
 Else
   Debug.Print "não existe arquivo de índices"
   ReDim linhas(0 To 1000)
   nroWp = 0
 End If
 SetTempoLimite tLim, duracao
 ti = Timer
 If bMostraProgress Then
   ProgBar.Show
   ProgBar.Indica 0, duracao, ""
 End If
 ReDim Preserve listaJPegsAvailable(0 To nroWp + 1000) ' cabem até mais 1000...
 Do
   bArquivoJaExiste = False
   If getLandscapeBackgroundFromWindowsSpotligh(DescrImg) Then
     With DescrImg
       ' vejo se arquivo já existe
       For i = 0 To nroWp - 1
         If listaSHA256(i) = .AssinaturaSHA256 Then
           bArquivoJaExiste = True
           Exit For
         End If
       Next i
       ' concateno o fim com valores aleatórios para evitar arquivos com o "mesmo nome"
       nomeArqImg = normaliza(.titulo)
       If bArquivoJaExiste Then
         Debug.Print nomeArqImg & " - Arquivo já existia!"
       Else
         nomeArqImg = nomeArqImg & "_" & .Width & "x" & .Height & "_" & Right$("000" & _
           Hex$(Int(64 * 1024& * Rnd)), 4) & Right$("000" & Hex$(Int(64 * 1024& * Rnd)), 4) & _
           Right$("000" & Hex$(Int(64 * 1024& * Rnd)), 4) & ".jpg"
         Debug.Print nomeArqImg
         imgString = GetURLAsString(.url)
         If Len(imgString) <> 0 Then
           If Dir$(pathFolderImg, vbDirectory) = "" Then MkDir pathFolderImg
           fullPathArqImg = pathFolderImg & "\" & nomeArqImg
           If Dir$(fullPathArqImg, vbNormal) <> "" Then Kill fullPathArqImg
           iLivre = FreeFile
           Open fullPathArqImg For Binary As #iLivre
           Put #iLivre, , imgString
           Close #iLivre
           imgString = ""
           linhas(nroWp) = .AssinaturaSHA256 & nomeArqImg
           listaJPegsAvailable(nroWp) = nomeArqImg
           listaSHA256(nroWp) = .AssinaturaSHA256
           nroWp = nroWp + 1
         End If
         
'         Set stdPic = getStdPictureFromURL(.url)
'         bSucesso = False
'         If Not stdPic Is Nothing Then
'           HandFreeImage = FreeImage_CreateFromOlePicture(stdPic)
'           If HandFreeImage <> 0 Then
'             If FreeImage_Save(FIF_JPEG, HandFreeImage, pathFolderImg & "\" & nomeArqImg, FISO_JPEG_QUALITYSUPERB) Then
''               SavePicture stdPic, pathFolderImg & "\" & nomeArqImg
'               bSucesso = True
'               linhas(nroWp) = .AssinaturaSHA256 & nomeArqImg
'               listaSHA256(nroWp) = .AssinaturaSHA256
'               nroWp = nroWp + 1
'             End If
'             FreeImage_UnloadEx (HandFreeImage)
'           End If
'         Else
'           Debug.Print "Falha criar imagem..."
'         End If
       End If
     End With
   Else
     Pausa 0.1
   End If
   If bMostraProgress Then ProgBar.Indica CorrigeTimerDif(ti, Timer), duracao, ""
   DoEvents
 Loop While Not TestaTempoLimiteAtingidoTimer(tLim)
 If bMostraProgress Then
   With ProgBar
    .Indica 100, 100, "[" & nroWp & "] wallpapers available!"
'   MsgboxAssincrono "[" & nroWp & "] wallpapers available!", App.EXEName, 3, vbOKOnly Or vbApplicationModal
   End With
   Pausa 2.5
   Unload ProgBar
 End If

 If nroWp = 0 Then
   MsgBox "Sorry. As I could not download any Wallpaper picture from the Internet, I can not operate. Please try later and garantee that this program has Internet access...", vbCritical Or vbOKOnly Or vbApplicationModal
   End
 End If
 ReDim Preserve listaJPegsAvailable(0 To nroWp - 1)
 Erase listaSHA256
 ReDim Preserve linhas(0 To nroWp - 1)
 If gravaCfgHash(pathArqIdxWallp & ".tmp", linhas, "chvIndice") Then
   If Dir$(pathArqIdxWallp) <> "" Then
     SetAttr pathArqIdxWallp, vbNormal
     Kill pathArqIdxWallp
   End If
   Name pathArqIdxWallp & ".tmp" As pathArqIdxWallp
 End If
 acertaNroImgsNoMenu
 descarregaWallpapers = nroWp
 menuCmds.Enabled = True
 Exit Function
trataerro:
  menuCmds.Enabled = True
End Function

Private Sub RegeraIDXFromFiles(pathNameIdxGerado As String)
  Dim ti As Single
  Dim b64 As New Base64
  Dim dadosS As String
  Dim enc As New CriptoHash
  Dim hashHex As String
  Dim fso As New FileSystemObject
  Dim arqs As Files, umArq As File
  Dim linhas() As String
  Dim n As Long
  
  ti = Timer
  If Dir$(pathNameIdxGerado, vbNormal) <> "" Then Kill pathNameIdxGerado
  ReDim linhas(0 To 8000)
  Set arqs = fso.GetFolder(pathFolderImg).Files
  n = 0
  For Each umArq In arqs
    If Right$(umArq.name, 4) = ".jpg" Then
      hashHex = enc.HashFile(umArq.path, SHA_256)
      linhas(n) = b64.Encode(b64.cadHex2cadBin(hashHex)) & umArq.name
'      Debug.Print linhas(n)
      n = n + 1
    End If
  Next umArq
  ReDim Preserve linhas(0 To n - 1)
  Debug.Print Timer - ti
  gravaCfgHash pathNameIdxGerado, linhas, "chvIndice"
End Sub


Private Sub Form_Load()

  Dim n As Long, j As Long, trabS As String
'  pathDebug = appPath & "debug.txt"
'  If Dir$(pathDebug, vbNormal) <> "" Then Kill pathDebug
  Dim currentWallpFileName As String, idxWalpAtual As Long

  With TrayIcon
    .cbSize = Len(TrayIcon)
    .hwnd = Me.hwnd
    .uID = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = "Periodically changes the DeskTop Wallpaper" & vbNullChar
  End With
  bColocouIconNaTray = False
'  geralogDebug "1"
  preparaCorMenus Me.hwnd
'  geralogDebug "2"
  
  appDataPath = Environ("appdata")
  If Right$(appDataPath, 1) <> "\" Then appDataPath = appDataPath & "\"
  appDataPath = appDataPath & App.EXEName & "\"
  If Dir$(appDataPath, vbDirectory) = "" Then MkDir appDataPath
'  pathFolderImg = Environ("userprofile") & "\Pictures\" & App.EXEName ' SEM BARRA NO FINAL!!!!
  pathFolderImg = Environ("Public") & "\Pictures\" & App.EXEName ' SEM BARRA NO FINAL!!!!
  If Dir$(pathFolderImg, vbDirectory) = "" Then MkDir pathFolderImg
  pathArqIdxWallp = pathFolderImg & "\index.idx"

  If Dir$(pathFolderImg, vbDirectory) = "" Then MkDir pathFolderImg
  pathCfg = appDataPath & App.EXEName & ".cfg"
  If Not leCfgHash(pathCfg, linhasCfg, "chvConfig") Then
    If Not leCfgHash(appPath & App.EXEName & ".cfg", linhasCfg, "chvConfig") Then
'      End
     ' Intervalo/livre_futuro/livre_futuro/livre_futuro/livre_futuro/livre_futuro/livre_futuro
      ReDim linhasCfg(0 To 7)
      linhasCfg(0) = "20190206:Developed by Mário Cesar Gomez Segura 'just for fun' as FREEWARE."
      intervaloTrocas = 4 * 3600& ' 4 horas
      linhasCfg(1) = intervaloTrocas
'      linhasCfg(2) = 0
      salvaCFG
    End If
  Else
'  geralogDebug "3"
  End If
  intervaloTrocas = CLng(linhasCfg(1))
'  If WillRunAtStartup(App.EXEName, valor) Then
  If WillRunAtStartupAlternative(App.EXEName) Then
'    cmdExe = appPath & App.EXEName & ".exe"
'    mRunStartup.Checked = (LCase$(Left$(Replace(valor, """", ""), Len(cmdExe))) = LCase$(cmdExe))
     mRunStartup.Checked = True
  Else
    mRunStartup.Checked = False
  End If
'  geralogDebug "4"
'  If GetTipoConexaoWeb = TCI_NOTAVAILABLE Then
'    TimerTentaDescarregar.Enabled = true
'  Else
'    ' descarregaImagens
'    nWp = descarregaWallpapers(TEMPODESCARGARAPIDA, False)
'    If nWp = 0 Then
'        nWp = descarregaWallpapers(TEMPODESCARGARAPIDA, False)
'        If nWp = 0 Then End
'    End If
'  End If
  Rnd -1
  Randomize
  GetCursorPos lastMousePos

  If Not pegaListaJPegsAvailable() Then End
  currentWallpFileName = getCurrentWallpPath
  currentWallpFileName = Mid$(currentWallpFileName, InStrRev(currentWallpFileName, "\") + 1)
  idxWalpAtual = -1
  For n = 0 To nroWp - 1
    If listaJPegsAvailable(n) = currentWallpFileName Then
      idxWalpAtual = n
      Exit For
    End If
  Next n
  If idxWalpAtual = -1 Then 'imagem atual não extá na lista
   ' embaralho lista
    For n = 0 To nroWp - 1
      j = Int(Rnd() * nroWp)
      trabS = listaJPegsAvailable(n)
      listaJPegsAvailable(n) = listaJPegsAvailable(j)
      listaJPegsAvailable(j) = trabS
    Next n
  Else ' deixo imagem atual no final da lista, e embaralho o resto da lista
    listaJPegsAvailable(idxWalpAtual) = listaJPegsAvailable(nroWp - 1)
    listaJPegsAvailable(nroWp - 1) = currentWallpFileName
    For n = 0 To nroWp - 2
      j = Int(Rnd() * (nroWp - 1))
      trabS = listaJPegsAvailable(n)
      listaJPegsAvailable(n) = listaJPegsAvailable(j)
      listaJPegsAvailable(j) = trabS
    Next n
  End If
'
  iproxImg = 0
  Load FormBkg
  pegaProxArqImg False
  FormBkg.CarregaImagemFundo imagemBase
  toEscuro = 0
  Brightness = 0
  setIntervaloTroca intervaloTrocas
'  Me.Hide
  Timer1.Enabled = True
  divTimerTentaDescarregar = 1 ' para executar em 1 segundo!!!!!
  With TimerTentaDescarregar
    .Interval = 1000 ' 1 segundo
    .Enabled = True
  End With
'  RegeraIDXFromFiles appPath & "idxRegerado.idx"
  
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static bRunning As Boolean
Dim cmd As Long
 If bRunning Then Exit Sub
 bRunning = True
 On Error Resume Next
 cmd = x / Screen.TwipsPerPixelX
 Select Case cmd
   Case WM_RBUTTONDOWN
      SetForegroundWindow Me.hwnd
      PopupMenu menuCmds, , , , mNomeLocal
   Case WM_LBUTTONDOWN
'      Debug.Print Shift
      If Shift <> 0 Then
        If iproxImg = 0 Then
          iproxImg = nroWp - 2
        Else
          If iproxImg = 1 Then
            iproxImg = nroWp - 1
          Else
            iproxImg = iproxImg - 2
          End If
        End If
      End If
      Me.Enabled = False
      pegaProxArqImg True
      Me.Enabled = True
 End Select
 bRunning = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim umF As Form
  Timer1.Enabled = False
  TimerIntervalo.Enabled = False
  TimerTampering.Enabled = False
  TimerTentaDescarregar.Enabled = False
  DoEvents
'  geralogDebug "200"
  If bColocouIconNaTray Then
    Shell_NotifyIcon NIM_DELETE, TrayIcon
  End If
'  geralogDebug "201"
  For Each umF In Forms
    If Not umF Is Me Then Unload umF
  Next umF
'  geralogDebug "202"
End Sub

Private Sub mAbout_Click()
  frmAbout.Show
End Sub

Private Sub mDownloadWpapers_Click()
  Dim nWp As Long
  nWp = descarregaWallpapers(TEMPODESCARGALONGA, True)
End Sub

Private Sub mExit_Click()
  Unload Me
End Sub
Private Sub salvaCFG()
  gravaCfgHash pathCfg, linhasCfg, "chvConfig"
End Sub

Private Sub mInterval_Click()
  Dim novoIntervaloS As String, novoIntervalo As Long
  Do
    novoIntervaloS = InputBox("Define interval in seconds(>=20 and <= 14400 [4h], 0 to disable):", "Interval between changes", intervalo)
    If novoIntervaloS = "" Then Exit Sub
    If IsNumeric(novoIntervaloS) Then
      novoIntervalo = Val(novoIntervaloS)
      If (novoIntervalo >= 20 And novoIntervalo <= 14400) Or (novoIntervalo = 0) Then Exit Do
    End If
  Loop
  setIntervaloTroca novoIntervalo
  salvaCFG
End Sub

Private Sub mKeywords_Click()
 Dim cmd As String
 cmd = "cmd /c start """"  ""https://www.google.com/search?q=" & Replace$(Replace$(localDaImagemStr, "__", "_"), "_", "+") & """"
 Shell cmd
End Sub

Private Sub mOpenPicFolder_Click()
  Dim cmd As String
  If Dir$(pathFolderImg, vbDirectory) = "" Then MkDir pathFolderImg
  cmd = "cmd /c start """" """ & pathFolderImg & """"
  Shell cmd ', vbHide

End Sub

Private Sub mRunStartup_Click()
  mRunStartup.Checked = Not mRunStartup.Checked
  If mRunStartup.Checked Then
'    linhasCfg(2) = 1
'    SetRunAtStartup App.EXEName, ".exe", appPath, True, CADINDSTARTUP
    SetRunAtStartupAlternative App.EXEName, ".exe", appPath, True, CADINDSTARTUP
  Else    ' se não recriar a chave, em cada ativação fica tentando 'reinstalar' a chave!!!!!
'    linhasCfg(2) = 0
'    SetRunAtStartup App.EXEName, "", "desativada", True
     
     '!!!!!!!!!! no caso da alternativa, isto não acontece!!!
    SetRunAtStartupAlternative App.EXEName, ".exe", appPath, False

  End If
  salvaCFG

End Sub

Private Sub Timer1_Timer()
  Dim mousePos As POINTAPI
'  Dim ti As Single
  Dim handleWindowSobMouse As Long
  Dim bAjustaBrightness As Boolean
'  Dim stdP As StdPicture
  Static div10 As Byte
  If Not bColocouIconNaTray Then
    If Shell_NotifyIcon(NIM_ADD, TrayIcon) <> 0 Then bColocouIconNaTray = True
  End If
'  If GetCursorPos(mousePos) Then
'    handleWindowSobMouse = WindowFromPoint(mousePos.x, mousePos.y)
'    If handleWindowSobMouse = handleJanelaMaisAtrasVistaPeloMouse Then
'      With mousePos
'        If Abs(.x - lastMousePos.x) + Abs(.y - lastMousePos.y) > 10 Then
'          bescurecendo = True
'          toEscuro = TOESCUROBASE
'        End If
'      End With
'    Else
'      If toEscuro > 20 Then toEscuro = 20
'    End If
'    If toEscuro > 0 Then
'      toEscuro = toEscuro - 1
'      If toEscuro = 0 Then bescurecendo = False
'    End If
'    lastMousePos = mousePos
'  End If
  If div10 = 0 Then
    If GetCursorPos(mousePos) Then
      With mousePos
        If Abs(.x - lastMousePos.x) + Abs(.y - lastMousePos.y) > 10 Then ' moveu
          handleWindowSobMouse = WindowFromPoint(.x, .y)
          If handleWindowSobMouse = handleJanelaMaisAtrasVistaPeloMouse Then
            bescurecendo = True
            toEscuro = TOESCUROBASE
          Else
            If toEscuro > 20 Then toEscuro = 20
          End If
        End If
      End With
      lastMousePos = mousePos
    End If
  End If
  div10 = (div10 + 1) Mod 10
  
  If toEscuro > 0 Then
    toEscuro = toEscuro - 1
    If toEscuro = 0 Then bescurecendo = False
  End If
  
  If bescurecendo Then
    If Brightness > MENORBRIGHTNESS Then
      Brightness = Brightness - DELTABRIGHTNESS
      If Brightness < MENORBRIGHTNESS Then Brightness = MENORBRIGHTNESS
'       Debug.Print Brightness
      ' posso deixar mais escuro
      If bAjustaBrightness = False Then
        Muda_prioridade REALTIME_PRIORITY_CLASS, THREAD_PRIORITY_TIME_CRITICAL
        bAjustaBrightness = True
'        Debug.Print "Prioridade Alta"
      End If
    Else ' já chegou no mínimo-> nada faz
      If bAjustaBrightness Then
        Muda_prioridade NORMAL_PRIORITY_CLASS, THREAD_PRIORITY_IDLE
        bAjustaBrightness = False
'        Debug.Print "Prioridade Normal"
      End If
    End If
  Else
    If Brightness < 0 Then
       Brightness = Brightness + DELTABRIGHTNESS
'       Debug.Print Brightness
       If Brightness > 0 Then Brightness = 0
       If bAjustaBrightness = False Then
         Muda_prioridade REALTIME_PRIORITY_CLASS, THREAD_PRIORITY_TIME_CRITICAL
         bAjustaBrightness = True
'         Debug.Print "Prioridade Alta"
       End If
      ' posso deixar mais claro
    Else ' já está no máximo-> nada faz
      If bAjustaBrightness Then
        Muda_prioridade NORMAL_PRIORITY_CLASS, THREAD_PRIORITY_IDLE
        bAjustaBrightness = False
'        Debug.Print "Prioridade Normal"
      End If
    End If
  End If
  If bAjustaBrightness Then
    'Brightness: -100 a 100
    'Bright: 0 a 2
    BrightVet VetIn(), VetOut(), pWidth * pHeight, ((Brightness + 100) / 100#)
    
    SetDIBitsToDevice FormBkg.hdc, 0, 0, pWidth, pHeight, 0, 0, 0, pHeight, VetOut(0), bi32BitInfo, 0
    FormBkg.Refresh
'      Debug.Print (Timer - ti) * 1000 & " ms"
  End If
End Sub

Private Sub timerIntervalo_Timer()
    
  If intervCnt <> 0 Then
    intervCnt = intervCnt - 1
'    Debug.Print intervCnt
  Else
    pegaProxArqImg True
  End If
End Sub

Private Sub TimerTampering_Timer()
  Static TO_20 As Integer
  If (Not TestaSeOperacional) Then
    TimerTampering.Enabled = False
'      geralogDebug "700"
    Unload Me
    Exit Sub
  End If
  If Not btestouTampering Then
    If TO_20 >= 20 Then
      btestouTampering = True
      If Not bInsideIDE Then
        If Not CheckHashConcatenadoEXE(appPath & App.EXEName & ".exe") Then
          bEncerrandoPrograma = True
          MsgboxAssincrono "!!! This program has been tampered with. !!!" & vbCrLf & _
                           " Please retrieve an original copy of it." & vbCrLf & _
                           " You'd better also test your system for viruses.", "Warning!!!", 10, _
                           vbCritical Or vbOKOnly Or vbMsgBoxSetForeground Or vbSystemModal
          Unload Form1
          End
        End If
      End If
    Else
      TO_20 = TO_20 + 1
    End If
  End If
End Sub

Private Sub TimerTentaDescarregar_Timer()
  Dim nWp As Long
  divTimerTentaDescarregar = divTimerTentaDescarregar - 1
  If divTimerTentaDescarregar <= 0 Then
    If GetTipoConexaoWeb <> TCI_NOTAVAILABLE Then
      divTimerTentaDescarregar = 3600 ' à cada 1 hora
      TimerTentaDescarregar.Enabled = False
      ' descarregaImagens
'      Debug.Print "Descarregando"
      nWp = descarregaWallpapers(TEMPODESCARGARAPIDA, False)
'      Debug.Print "Acabou"
      TimerTentaDescarregar.Enabled = True
    Else
      divTimerTentaDescarregar = 10 ' para provocar retry por insucesso em 10 segundos
    End If
  End If
End Sub
