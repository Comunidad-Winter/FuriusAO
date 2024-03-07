Attribute VB_Name = "wskapiAO"
Option Explicit

'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then
#Const WSAPI_CREAR_LABEL = True

Private Const SD_RECEIVE As Long = &H0
Private Const SD_SEND As Long = &H1
Private Const SD_BOTH As Long = &H2

Private Const MAX_TIEMPOIDLE_COLALLENA = 1
Private Const MAX_COLASALIDA_COUNT = 800

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'FIXIT: As Any no se admite en Visual Basic .NET. Utilice un tipo específico.              FixIT90210ae-R5608-H1984
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WS_CHILD = &H40000000
Public Const GWL_WNDPROC = (-4)

Private Const SIZE_RCVBUF As Long = 8192
Private Const SIZE_SNDBUF As Long = 8192

Public Type tSockCache
    Sock As Long
    Slot As Long
End Type

Public WSAPISock2Usr As New Collection

Public OldWProc As Long
Public ActualWProc As Long
Public hWndMsg As Long
Public SockListen As Long
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
Public Sub IniciaWsApi()
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then

hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", 0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0&)
OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

Dim Desc As String
Call StartWinsock(Desc)

'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
End Sub
Public Function StartWinsock(sDescription As String) As Boolean

Dim StartupData As WSADataType
If Not WSAStartedUp Then
    If Not WSAStartup(&H101, StartupData) Then
        WSAStartedUp = True
        sDescription = StartupData.szDescription
    Else: WSAStartedUp = False
    End If
End If
StartWinsock = WSAStartedUp

End Function
Public Sub LimpiaWsApi(ByVal hwnd As Long)
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then

If WSAStartedUp Then
    Call EndWinsock
End If

'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
End Sub
Public Function BuscaSlotSock(S As Long) As Long
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).ConnID = S Then
        BuscaSlotSock = i
        Exit Function
    End If
Next i

BuscaSlotSock = -1

'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
End Function
Public Function WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then

Dim Ret As Long
Dim Tmp As String

WndProc = 0

Select Case msg
Case 1025
    Dim S As Long, E As Long
    Dim N As Integer
    S = wParam
    E = WSAGetSelectEvent(lParam)
    
    Select Case E
    Case FD_ACCEPT
        Call EventoSockAccept(S)
    Case FD_READ
        
        N = BuscaSlotSock(S)
        If N < 0 Then
            Call ApiCloseSocket(S)
            Exit Function
        End If
    
    
        Tmp = Space(4096)
        
        Ret = recv(S, Tmp, Len(Tmp), 0)
        
        If Ret < 0 Then
            Call CloseSocket(N)
            Exit Function
        End If
        
'FIXIT: Reexmplazar la función 'Left' con la función 'Left$'.                               FixIT90210ae-R9757-R1B8ZE
        Tmp = Left(Tmp, Ret)
        
        Call EventoSockRead(N, Tmp)
        
    Case FD_CLOSE
        N = BuscaSlotSock(S)
                
        If N < 0 Then
            Call ApiCloseSocket(S)
        Else: Call EventoSockClose(N)
        End If
        
    End Select
Case Else
    WndProc = CallWindowProc(OldWProc, hwnd, msg, wParam, lParam)
End Select

'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
End Function
Public Sub WsApiEnviar(Slot As Integer, str As String)
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then
Dim Ret As String

If UserList(Slot).ConnID > -1 Then
    Ret = send(ByVal UserList(Slot).ConnID, ByVal str, ByVal Len(str), ByVal 0)
    If Ret < 0 Then Exit Sub
End If
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If

End Sub
Public Sub LogCustom(ByVal str As String)
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If (UsarQueSocket = 1) Then

Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\logs\custom.log" For Append Shared As #nfile
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, Date & " " & Time & " " & str
Close #nfile

Exit Sub

errhandler:

'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
End Sub

Public Sub EventoSockAccept(ByVal SockID As Long)
#If UsarQueSocket = 1 Then
'==========================================================
'USO DE LA API DE WINSOCK
'========================
    
    Dim NewIndex As Integer
    Dim Ret As Long
    Dim Tam As Long, sa As sockaddr
    Dim NuevoSock As Long
    Dim i As Long
    Dim tStr As String
    
    Tam = sockaddr_size
    
    '=============================================
    'SockID es en este caso es el socket de escucha,
    'a diferencia de socketwrench que es el nuevo
    'socket de la nueva conn
    
'Modificado por Maraxus
    'Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 0)
    Ret = accept(SockID, sa, Tam)

    If Ret = INVALID_SOCKET Then
        i = Err.LastDllError
        Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
        Exit Sub
    End If
    
'    If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
'        Call WSApiCloseSocket(NuevoSock)
'        Exit Sub
'    End If

    NuevoSock = Ret
    
    'Seteamos el tamaño del buffer de entrada a 512 bytes
    If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de entrada " & i & ": " & GetWSAErrorString(i))
    End If
    'Seteamos el tamaño del buffer de salida a 1 Kb
    If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & GetWSAErrorString(i))
    End If

           
        


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   BIENVENIDO AL SERVIDOR!!!!!!!!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
    NewIndex = NextOpenUser ' Nuevo indice
     
    
    If NewIndex <= MaxUsers Then
        
        UserList(NewIndex).ip = GetAscIP(sa.sin_addr)
        'Busca si esta banneada la ip
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = UserList(NewIndex).ip Then
                'Call apiclosesocket(NuevoSock)
                tStr = "ERRSu IP se encuentra bloqueada en este servidor." & ENDC
                Call send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
                'Call SecurityIp.IpRestarConexion(sa.sin_addr)
                Call WSApiCloseSocket(NuevoSock)
                Exit Sub
            End If
        Next i
        
        
'    If False Then
        If aDos.MaxConexiones(UserList(NewIndex).ip) Then
            UserList(NewIndex).ConnID = -1
            Call aDos.RestarConexion(UserList(NewIndex).ip)
            'tStr = "ERRLimite de conexiones para su IP alcanzado." & ENDC
            'Call send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
            Call ApiCloseSocket(NuevoSock)
         'If SecurityIp.IPSecuritySuperaLimiteConexiones(sa.sin_addr) Then
         'Call WSApiCloseSocket(NuevoSock)
         Exit Sub
        End If
'    End If


        
        If NewIndex > LastUser Then LastUser = NewIndex
        
        UserList(NewIndex).SockPuedoEnviar = True
        UserList(NewIndex).ConnID = NuevoSock
        UserList(NewIndex).ConnIDvalida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        Set UserList(NewIndex).ColaSalida = New Collection
        
        Call AgregaSlotSock(NuevoSock, NewIndex)
    Else
        tStr = "ERRServer lleno." & ENDC
        Dim AAA As Long
        AAA = send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
        'Call SecurityIp.IpRestarConexion(sa.sin_addr)
        Call WSApiCloseSocket(NuevoSock)
    End If
    
#End If
End Sub

Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal Slot As Long)
Debug.Print "AgregaSockSlot"
#If (UsarQueSocket = 1) Then
If WSAPISock2Usr.Count > MaxUsers Then
    Call CloseSocket(Slot)
    Exit Sub
End If
WSAPISock2Usr.Add CStr(Slot), CStr(Sock)
#End If
End Sub



Public Sub EventoSockAccept2(SockID As Long)
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then
On Error GoTo ErrorHandler
    Dim NewIndex As Integer
    Dim Ret As Long
    Dim Tam As Long, sa As sockaddr
    Dim NuevoSock As Long
    Dim i As Long
        
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        
        Tam = sockaddr_size
        
        Ret = accept(SockID, sa, Tam)
        If Ret = INVALID_SOCKET Then
            Call LogCriticEvent("Error en Accept() API")
            Exit Sub
        End If
        NuevoSock = Ret
        
        UserList(NewIndex).ip = GetAscIP(sa.sin_addr)
        
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = UserList(NewIndex).ip Then
                Call ApiCloseSocket(NuevoSock)
                Exit Sub
            End If
        Next i
        
        '1 POR IP
       ' Dim rr As Integer
       ' For rr = 1 To LastUser
       '     If UserList(rr).ip = UserList(NewIndex).ip Then
       '         If rr <> NewIndex Then
       '             Call ApiCloseSocket(NuevoSock)
       '             Exit Sub
       '         End If
       '     End If
       ' Next rr
        '1 POR IP
        
        If aDos.MaxConexiones(UserList(NewIndex).ip) Then
            UserList(NewIndex).ConnID = -1
            Call aDos.RestarConexion(UserList(NewIndex).ip)
            Call ApiCloseSocket(NuevoSock)
        End If
        
        UserList(NewIndex).ConnID = NuevoSock
        Set UserList(NewIndex).CommandsBuffer = New CColaArray

    Else
        Call LogCriticEvent("No acepte conexion porque no tenia slots")
    End If
    Exit Sub
ErrorHandler:
    Call LogError("Error en EventoSockAccept. " & Err.Description)
    Call CloseSocket(NewIndex)
    
    
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
End Sub
Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos As String)
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then
On Error GoTo ErrorHandler
Dim T() As String
Dim LoopC As Long

UserList(Slot).RDBuffer = UserList(Slot).RDBuffer & Datos

If InStr(1, UserList(Slot).RDBuffer, Chr(2)) > 0 Then UserList(Slot).RDBuffer = "CLIENTEVIEJO" & ENDC

T = Split(UserList(Slot).RDBuffer, ENDC)
If UBound(T) > 0 Then
    UserList(Slot).RDBuffer = T(UBound(T))
    
    For LoopC = 0 To UBound(T) - 1
        If ClientsCommandsQueue = 1 Then
            If Len(T(LoopC)) > 0 Then If Not UserList(Slot).CommandsBuffer.Push(T(LoopC)) Then Call Cerrar_Usuario(Slot)
        Else
            If UserList(Slot).ConnID <> -1 Then
                Call HandleData(Slot, T(LoopC))
            Else: Exit Sub
            End If
        End If
    Next LoopC
End If

Exit Sub

ErrorHandler:
    Call LogError("Error en Socket read. " & Err.Description)
    Call CloseSocket(Slot)

'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
End Sub
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If Win16 Then
'FIXIT: Declare 'vMessage' con un tipo de datos de enlace en tiempo de compilación         FixIT90210ae-R1672-R1B8ZE
Public Function kSendData(ByVal S%, vMessage As Variant) As Integer
'FIXIT: '#ElseIf' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#ElseIf Win32 Then
'FIXIT: Declare 'vMessage' con un tipo de datos de enlace en tiempo de compilación         FixIT90210ae-R1672-R1B8ZE
Public Function kSendData(ByVal S&, vMessage As Variant) As Long
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
    Dim TheMsg() As Byte, sTemp$
    TheMsg = ""
    Select Case VarType(vMessage)
        Case 8209
            sTemp = vMessage
            TheMsg = sTemp
        Case 8
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
            #If Win32 Then
                sTemp = StrConv(vMessage, vbFromUnicode)
'FIXIT: '#Else' no se actualiza de forma fiable a Visual Basic .NET                        FixIT90210ae-R2789-H1984
            #Else
                sTemp = vMessage
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
            #End If
        Case Else
            sTemp = CStr(vMessage)
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
            #If Win32 Then
                sTemp = StrConv(vMessage, vbFromUnicode)
'FIXIT: '#Else' no se actualiza de forma fiable a Visual Basic .NET                        FixIT90210ae-R2789-H1984
            #Else
                sTemp = vMessage
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
            #End If
    End Select
    TheMsg = sTemp
    If UBound(TheMsg) > -1 Then
        kSendData = send(S, TheMsg(0), UBound(TheMsg) + 1, 0)
    End If
End Function
Public Sub EventoSockClose(Slot As Integer)

'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then
    If UserList(Slot).flags.UserLogged Then
        Call Cerrar_Usuario(Slot)
    Else: Call CloseSocket(Slot)
    End If
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If

End Sub


Public Sub WSApiReiniciarSockets()
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then
Dim i As Long
    
    If SockListen >= 0 Then Call ApiCloseSocket(SockListen)

    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And UserList(i).ConnIDvalida Then
            Call CloseSocket(i)
        End If
    Next i
    
    
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
    ReDim UserList(1 To MaxUsers + 10)
    For i = 1 To MaxUsers
        UserList(i).ConnID = -1
        UserList(i).ConnIDvalida = False
    Next i
    
    LastUser = 1
    NumUsers = 0
    
    Call LimpiaWsApi(frmMain.hwnd)
    Call Sleep(100)
    Call IniciaWsApi
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
End Sub

Public Sub WSApiCloseSocket(ByVal Socket As Long)
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then
Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
Call ShutDown(Socket, SD_BOTH)
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
End Sub
