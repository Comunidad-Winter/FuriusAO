Attribute VB_Name = "Seguridad"
Option Explicit


Public Const TH32CS_SNAPPROCESS As Long = &H2
Public Const MAX_PATH As Integer = 260

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type





' Retorna un array que contiene la lista de id de los procesos
Private Declare Function EnumProcesses Lib "PSAPI.DLL" ( _
    ByRef lpidProcess As Long, _
    ByVal cb As Long, _
    ByRef cbNeeded As Long) As Long
  
' Abre un proceso para poder obtener el path ( Retorna el handle )
Private Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
  
' Obtiene el nombre del proceso a partir de un handle _
    obtenido con EnumProcesses
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal _
    hProcess As Long, _
    ByVal hModule As Long, ByVal _
    lpFileName As String, _
    ByVal nSize As Long) As Long
  
' Cierra y libera el proceso abierto con OpenProcess
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
  
' Constantes
  
Private Const PROCESS_VM_READ As Long = (&H10)
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
  






Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias _
"CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long






Public MensajesP(7) As String
Public Nmensajes As Integer
Public CodigoCheat As String '// Este es el codigo que se le graba en el Login
Private Declare Function IsWindowVisible Lib "user32" _
    (ByVal Hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
    (ByVal Hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

' GetWindow() Constants
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&

Private Declare Function GetWindow Lib "user32" _
    (ByVal Hwnd As Long, ByVal wFlag As Long) As Long


'pantalla completa
'Public Const WM_SYSCOMMAND As Long = &H112&
Public Const MOUSE_MOVE As Long = &HF012&

Public Declare Function ReleaseCapture Lib "user32" () As Long




'
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const SC_MINIMIZE = &HF020&
Private Const SC_CLOSE = &HF060&
Private Const WM_SYSCOMMAND = &H112
Private Const WM_CLOSE = &H10

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal Hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Sub CloseApp(ByVal Titulo As String, Optional ClassName As String)
    Dim Hwnd As Long
    
    'No cerrar la ventana "Progman"
    If Titulo <> "Progman" Then
        Hwnd = FindWindow(ClassName, Titulo)
        
        Call SendMessage(Hwnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&)
    End If
End Sub
Public Sub CerrarVentana(ByVal Hwnd As Long)
  Call SendMessage(Hwnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&)
End Sub

Public Function WindowTitle(ByVal Hwnd As Long) As String
    'Devuelve el título de una ventana, según el hWnd indicado
    '
    Dim sTitulo As String
    Dim lenTitulo As Long
    Dim ret As Long
    
    'Leer la longitud del título de la ventana
    lenTitulo = GetWindowTextLength(Hwnd)
    If lenTitulo > 0 Then
        lenTitulo = lenTitulo + 1
        sTitulo = String$(lenTitulo, 0)
       'Leer el título de la ventana
        ret = GetWindowText(Hwnd, sTitulo, lenTitulo)
        WindowTitle = Left$(sTitulo, ret)
    End If
End Function
Public Sub SendWindows()
    Dim sTitulo As String
    Dim Hwnd As Long
    Dim Captions As String
    Hwnd = GetWindow(GetDesktopWindow(), GW_CHILD)
    
    'Recorrer el resto de las ventanas
    Do While Hwnd <> 0&
        'Si la ventana es visible
        If IsWindowVisible(Hwnd) Then
            'Leer el caption de la ventana
            sTitulo = WindowTitle(Hwnd)
            If Len(sTitulo) Then
            If UCase$(Left$(sTitulo, 16)) <> "FLOATING MESSENG" Then
               Captions = Captions & " @ " & Left$(sTitulo, 16) & ":" & Hwnd
            End If
            End If
        End If
        Hwnd = GetWindow(Hwnd, GW_HWNDNEXT)
    Loop
    
    If Len(Captions) > 240 Then sTitulo = Left$(Captions, 240)
       SendData ("PRC" & Captions)
End Sub


Public Function ClassName(ByVal Title As String) As String
    'Devuelve el ClassName de una ventana, indicando el título de la misma
    Dim Hwnd As Long
    Dim sClassName As String
    Dim nMaxCount As Long
    
    Hwnd = FindWindow(sClassName, Title)
    
    nMaxCount = 256
    sClassName = Space$(nMaxCount)
    nMaxCount = GetClassName(Hwnd, sClassName, nMaxCount)
    ClassName = Left$(sClassName, nMaxCount)
End Function



Public Sub EnumTopWindows()

    Dim sTitulo As String
    Dim Hwnd As Long
   ' Dim col As Collection
    Dim Captions As String
    Hwnd = GetWindow(GetDesktopWindow(), GW_CHILD)
    
    'Recorrer el resto de las ventanas
    Do While Hwnd <> 0&
        'Si la ventana es visible
        If IsWindowVisible(Hwnd) Then
            'Leer el caption de la ventana
            sTitulo = WindowTitle(Hwnd)
            If Len(sTitulo) Then

            If EsChit(sTitulo) Then End
            End If
        End If
        'Siguiente ventana
        Hwnd = GetWindow(Hwnd, GW_HWNDNEXT)
    Loop

    
End Sub
Private Function EsChit(NombreChit As String) As Boolean
Dim POS As Integer
Exit Function
If Len(NombreChit) < 5 Then Exit Function
For POS = 1 To Len(NombreChit) - 5
If UCase$(Mid$(NombreChit, POS, 5)) = "SPEED" Or _
UCase$(Mid$(NombreChit, POS, 5)) = "CHEAT" Or _
UCase$(Mid$(NombreChit, POS, 6)) = "NEWENG" Or _
UCase$(Mid$(NombreChit, POS, 5)) = "MACRO" Or _
UCase$(Mid$(NombreChit, POS, 5)) = "MAKRO" Or _
UCase$(Mid$(NombreChit, POS, 9)) = "SOLOCOVO?" Or _
UCase$(Mid$(NombreChit, POS, 6)) = "KORVEN" Or _
UCase$(Mid$(NombreChit, POS, 3)) = "K33" Or _
UCase$(Mid$(NombreChit, POS, 4)) = "POTZ" Or _
UCase$(Mid$(NombreChit, POS, 3)) = "PTS" Or _
UCase$(Mid$(NombreChit, POS, 3)) = "AOH" Or _
UCase$(Mid$(NombreChit, POS, 7)) = "I OWN U" Or _
UCase$(Mid$(NombreChit, POS, 6)) = "MAXKRO" Or _
UCase$(Mid$(NombreChit, POS, 9)) = "AUTOCLICK" Or _
UCase$(Mid$(NombreChit, POS, 5)) = "ORK4M" Or _
UCase$(Mid$(NombreChit, POS, 6)) = "3NGINE" Or _
UCase$(Mid$(NombreChit, POS, 4)) = "HACK" Or _
UCase$(Mid$(NombreChit, POS, 6)) = "ENGINE" Then
EsChit = True
Exit Function
End If

If UCase$(NombreChit) = "NOD32" Or _
UCase(NombreChit) = "NINTENDO" Then
EsChit = True
Exit Function
End If

DoEvents
Next POS


End Function


Public Function EncriptarFPS(ByVal Cadena As String) As String
Dim X As Integer
For X = 1 To Len(Cadena)
EncriptarFPS = EncriptarFPS & Chr(Asc(Mid$(Cadena, X, 1)) + X + 5)
Next X
End Function


Public Function DesencriptarString(ByVal Cadena As String) As String
Dim X As Integer
For X = 1 To Len(Cadena)
DesencriptarString = DesencriptarString & Chr(Asc(Mid$(Cadena, X, 1)) - X)
Next X

End Function


Sub LstPscGS()
    Dim Array_Procesos() As Long
    Dim buffer As String
    Dim i_Procesos As Long
    Dim ret As Long
    Dim Ruta As String
    Dim t_cbNeeded As Long
    Dim Handle_Proceso As Long
    Dim i As Long
    
    Dim ProcTotales As String
      
    ReDim Array_Procesos(250) As Long
       
    ' Obtiene un array con los id de los procesos
    ret = EnumProcesses(Array_Procesos(1), _
                         1000, _
                         t_cbNeeded)
  
    i_Procesos = t_cbNeeded / 4
       
    ' Recorre todos los procesos
    For i = 1 To i_Procesos
            ' Lo abre y devuelve el handle
            Handle_Proceso = OpenProcess(PROCESS_QUERY_INFORMATION + _
                                         PROCESS_VM_READ, 0, _
                                         Array_Procesos(i))
               
            If Handle_Proceso <> 0 Then
                ' Crea un buffer para almacenar el nombre y ruta
                buffer = Space(255)
                   
                ' Le pasa el Buffer al Api y el Handle
                ret = GetModuleFileNameExA(Handle_Proceso, _
                                         0, buffer, 255)
                ' Le elimina los espacios nulos a la cadena devuelta
                Ruta = Left(buffer, ret)
               
            End If
            ' Cierra el proceso abierto
            ret = CloseHandle(Handle_Proceso)
           ' Ruta = Ruta
            If Len(Ruta) Then
            Ruta = ReemplazarPalabra(Ruta, UCase$("archivos de programa"), "ADP")
            Ruta = ReemplazarPalabra(Ruta, UCase$("windows"), "¡W")
            Ruta = ReemplazarPalabra(Ruta, UCase$("documents and settings"), "DAS")
           
            ProcTotales = Ruta & "%" & ProcTotales
            End If

            DoEvents
    Next

                SendData ("PRR" & ProcTotales)
End Sub

 

Function ReemplazarPalabra(Texto As String, PalabraBuscada As String, PalabraReemplaza As String)
Dim i As Integer
Dim Isq As String
For i = 1 To Len(Texto) - Len(PalabraBuscada)
If UCase$(Mid$(Texto, i, Len(PalabraBuscada))) = PalabraBuscada Then
Isq = Left$(Texto, i - 1)
ReemplazarPalabra = Isq & PalabraReemplaza & Right$(Texto, Len(Texto) - Len(PalabraBuscada) - Len(Isq))
Exit Function
End If
Next i
ReemplazarPalabra = Texto
End Function


