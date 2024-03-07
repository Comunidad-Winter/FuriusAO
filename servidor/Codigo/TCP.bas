Attribute VB_Name = "TCP"
Option Explicit

Public Const SOCKET_BUFFER_SIZE = 3072
Public Enpausa As Boolean

Public Const COMMAND_BUFFER_SIZE = 1000
Public EnTorneo As Byte

Public Const NingunArma = 2
Dim Response As String
Dim Start As Single, Tmr As Single


Public Const ToIndex = 0
Public Const ToAll = 1
Public Const ToMap = 2
Public Const ToPCArea = 3
Public Const ToNone = 4
Public Const ToAllButIndex = 5
Public Const ToMapButIndex = 6
Public Const ToGM = 7
Public Const ToNPCArea = 8
Public Const ToGuildMembers = 9
Public Const ToAdmins = 10
Public Const ToPCAreaButIndex = 11
Public Const ToMuertos = 12
Public Const ToPCAreaVivos = 13
Public Const ToNPCAreaG = 14
Public Const ToPCAreaButIndexG = 15
Public Const ToGMArea = 16
Public Const ToPCAreaG = 17
Public Const ToAlianza = 18
Public Const ToCaos = 19
Public Const ToParty = 20
Public Const ToMoreAdmins = 21
Public Const ToDiosesYclan = 22


'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 0 Then
Public Const INVALID_HANDLE = -1
Public Const CONTROL_ERRIGNORE = 0
Public Const CONTROL_ERRDISPLAY = 1



Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_DISCONNECT = 7
Public Const SOCKET_ABORT = 8


Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7


Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2


Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5


Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_GGP = 2
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256



Public Const INADDR_ANY = "0.0.0.0"
Public Const INADDR_LOOPBACK = "127.0.0.1"
Public Const INADDR_NONE = "255.055.255.255"


Public Const SOCKET_READ = 0
Public Const SOCKET_WRITE = 1
Public Const SOCKET_READWRITE = 2


Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1


Public Const WSABASEERR = 24000
Public Const WSAEINTR = 24004
Public Const WSAEBADF = 24009
Public Const WSAEACCES = 24013
Public Const WSAEFAULT = 24014
Public Const WSAEINVAL = 24022
Public Const WSAEMFILE = 24024
Public Const WSAEWOULDBLOCK = 24035
Public Const WSAEINPROGRESS = 24036
Public Const WSAEALREADY = 24037
Public Const WSAENOTSOCK = 24038
Public Const WSAEDESTADDRREQ = 24039
Public Const WSAEMSGSIZE = 24040
Public Const WSAEPROTOTYPE = 24041
Public Const WSAENOPROTOOPT = 24042
Public Const WSAEPROTONOSUPPORT = 24043
Public Const WSAESOCKTNOSUPPORT = 24044
Public Const WSAEOPNOTSUPP = 24045
Public Const WSAEPFNOSUPPORT = 24046
Public Const WSAEAFNOSUPPORT = 24047
Public Const WSAEADDRINUSE = 24048
Public Const WSAEADDRNOTAVAIL = 24049
Public Const WSAENETDOWN = 24050
Public Const WSAENETUNREACH = 24051
Public Const WSAENETRESET = 24052
Public Const WSAECONNABORTED = 24053
Public Const WSAECONNRESET = 24054
Public Const WSAENOBUFS = 24055
Public Const WSAEISCONN = 24056
Public Const WSAENOTCONN = 24057
Public Const WSAESHUTDOWN = 24058
Public Const WSAETOOMANYREFS = 24059
Public Const WSAETIMEDOUT = 24060
Public Const WSAECONNREFUSED = 24061
Public Const WSAELOOP = 24062
Public Const WSAENAMETOOLONG = 24063
Public Const WSAEHOSTDOWN = 24064
Public Const WSAEHOSTUNREACH = 24065
Public Const WSAENOTEMPTY = 24066
Public Const WSAEPROCLIM = 24067
Public Const WSAEUSERS = 24068
Public Const WSAEDQUOT = 24069
Public Const WSAESTALE = 24070
Public Const WSAEREMOTE = 24071
Public Const WSASYSNOTREADY = 24091
Public Const WSAVERNOTSUPPORTED = 24092
Public Const WSANOTINITIALISED = 24093
Public Const WSAHOST_NOT_FOUND = 25001
Public Const WSATRY_AGAIN = 25002
Public Const WSANO_RECOVERY = 25003
Public Const WSANO_DATA = 25004
Public Const WSANO_ADDRESS = 2500
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If

'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
Public Data(1 To 3, 1 To 2, 1 To 2, 1 To 2) As Double
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
Public Onlines(1 To 3) As Long

Public Const Minuto = 1
Public Const Hora = 2
Public Const Dia = 3

Public Const Actual = 1
Public Const Last = 2

Public Const Enviada = 1
Public Const Recibida = 2

Public Const Mensages = 1
Public Const Letras = 2

Sub DarCuerpoYCabeza(UserBody As Integer, UserHead As Integer, Raza As Byte, Gen As Byte)

Select Case Gen
   Case HOMBRE
        Select Case Raza
        
                Case HUMANO
                    UserHead = CInt(RandomNumber(1, 24))
                    If UserHead > 24 Then UserHead = 24
                    UserBody = 1
                Case ELFO
                    UserHead = CInt(RandomNumber(1, 7)) + 100
                    If UserHead > 107 Then UserHead = 107
                    UserBody = 2
                Case ELFO_OSCURO
                    UserHead = CInt(RandomNumber(1, 4)) + 200
                    If UserHead > 204 Then UserHead = 204
                    UserBody = 3
                Case ENANO
                    UserHead = RandomNumber(1, 4) + 300
                    If UserHead > 304 Then UserHead = 304
                    UserBody = 52
                Case GNOMO
                    UserHead = RandomNumber(1, 3) + 400
                    If UserHead > 403 Then UserHead = 403
                    UserBody = 52
                Case Else
                    UserHead = 1
                    UserBody = 1
            
        End Select
   Case MUJER
        Select Case Raza
                Case HUMANO
                    UserHead = CInt(RandomNumber(1, 4)) + 69
                    If UserHead > 73 Then UserHead = 73
                    UserBody = 1
                Case ELFO
                    UserHead = CInt(RandomNumber(1, 5)) + 169
                    If UserHead > 174 Then UserHead = 174
                    UserBody = 2
                Case ELFO_OSCURO
                    UserHead = CInt(RandomNumber(1, 5)) + 269
                    If UserHead > 274 Then UserHead = 274
                    UserBody = 3
                Case GNOMO
                    UserHead = RandomNumber(1, 4) + 469
                    If UserHead > 473 Then UserHead = 473
                    UserBody = 52
                Case ENANO
                    UserHead = RandomNumber(1, 3) + 369
                    If UserHead > 372 Then UserHead = 372
                    UserBody = 52
                Case Else
                    UserHead = 70
                    UserBody = 1
        End Select
End Select

   
End Sub
Sub ConnectNewUser(userindex As Integer, Name As String, Password As String, _
Body As Integer, Head As Integer, UserRaza As Byte, UserSexo As Byte, _
UA1 As String, UA2 As String, UA3 As String, UA4 As String, UA5 As String, _
US1 As String, US2 As String, US3 As String, US4 As String, US5 As String, _
US6 As String, US7 As String, US8 As String, US9 As String, US10 As String, _
US11 As String, US12 As String, US13 As String, US14 As String, US15 As String, _
US16 As String, US17 As String, US18 As String, US19 As String, US20 As String, _
US21 As String, US22 As String, UserEmail As String, Hogar As Byte, PCL As String, PasCod As String)

Dim i As Integer

    If PasCod = MD5String("estAnolasabesFdYl" & hex(10) & (UserList(userindex).flags.ValCoDe - 23)) Then
        UserList(userindex).flags.Devolvio = True
    Else
        CloseSocket (userindex)
        Exit Sub
    End If



If Restringido Then
    Call SendData(ToIndex, userindex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
    Exit Sub
End If

If Not NombrePermitido(Name) Then
    Call SendData(ToIndex, userindex, 0, "ERRLos nombres de los personajes deben pertencer a la fantasia, el nombre indicado es invalido.")
    Call SendData(ToIndex, userindex, 0, "V8V" & 2)
    Exit Sub
End If

If Not AsciiValidos(Name) Then
    Call SendData(ToIndex, userindex, 0, "ERRNombre invalido.")
    Call SendData(ToIndex, userindex, 0, "V8V" & 2)
    Exit Sub
End If



If Left$(Name, 1) = " " Then
    Call SendData(ToIndex, userindex, 0, "ERRNombre invalido.")
    Call SendData(ToIndex, userindex, 0, "V8V" & 2)
    Exit Sub
End If

Dim LoopC As Integer
Dim totalskpts As Long
  

If ExistePersonaje(Name) Then
    Call SendData(ToIndex, userindex, 0, "ERRYa existe el personaje.")
    Call SendData(ToIndex, userindex, 0, "V8V" & 2)
    Exit Sub
End If

UserList(userindex).flags.Muerto = 0
UserList(userindex).flags.Escondido = 0

UserList(userindex).Name = Name
UserList(userindex).Clase = CIUDADANO
UserList(userindex).Raza = UserRaza
UserList(userindex).Genero = UserSexo
UserList(userindex).Email = UserEmail
UserList(userindex).Hogar = Hogar

Select Case UserList(userindex).Raza
    Case HUMANO
        UserList(userindex).Stats.UserAtributosBackUP(fuerza) = UserList(userindex).Stats.UserAtributosBackUP(fuerza) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Constitucion) = UserList(userindex).Stats.UserAtributosBackUP(Constitucion) + 2
    Case ELFO
        UserList(userindex).Stats.UserAtributosBackUP(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad) + 3
        UserList(userindex).Stats.UserAtributosBackUP(Constitucion) = UserList(userindex).Stats.UserAtributosBackUP(Constitucion) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) = UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Carisma) = UserList(userindex).Stats.UserAtributosBackUP(Carisma) + 2
    Case ELFO_OSCURO
        UserList(userindex).Stats.UserAtributosBackUP(fuerza) = UserList(userindex).Stats.UserAtributosBackUP(fuerza) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Carisma) = UserList(userindex).Stats.UserAtributosBackUP(Carisma) - 3
        UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) = UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) + 2
    Case ENANO
        UserList(userindex).Stats.UserAtributosBackUP(fuerza) = UserList(userindex).Stats.UserAtributosBackUP(fuerza) + 3
        UserList(userindex).Stats.UserAtributosBackUP(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad) - 1
        UserList(userindex).Stats.UserAtributosBackUP(Constitucion) = UserList(userindex).Stats.UserAtributosBackUP(Constitucion) + 3
        UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) = UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) - 6
        UserList(userindex).Stats.UserAtributosBackUP(Carisma) = UserList(userindex).Stats.UserAtributosBackUP(Carisma) - 3
    Case GNOMO
        UserList(userindex).Stats.UserAtributosBackUP(fuerza) = UserList(userindex).Stats.UserAtributosBackUP(fuerza) - 5
        UserList(userindex).Stats.UserAtributosBackUP(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad) + 4
        UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) = UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) + 3
        UserList(userindex).Stats.UserAtributosBackUP(Carisma) = UserList(userindex).Stats.UserAtributosBackUP(Carisma) + 1
End Select

If Not ValidateAtrib(userindex) Then
    Call SendData(ToIndex, userindex, 0, "ERRAtributos invalidos.")
    Call SendData(ToIndex, userindex, 0, "V8V" & 2)
    Exit Sub
End If

UserList(userindex).Stats.UserSkills(1) = val(US1)
UserList(userindex).Stats.UserSkills(2) = val(US2)
UserList(userindex).Stats.UserSkills(3) = val(US3)
UserList(userindex).Stats.UserSkills(4) = val(US4)
UserList(userindex).Stats.UserSkills(5) = val(US5)
UserList(userindex).Stats.UserSkills(6) = val(US6)
UserList(userindex).Stats.UserSkills(7) = val(US7)
UserList(userindex).Stats.UserSkills(8) = val(US8)
UserList(userindex).Stats.UserSkills(9) = val(US9)
UserList(userindex).Stats.UserSkills(10) = val(US10)
UserList(userindex).Stats.UserSkills(11) = val(US11)
UserList(userindex).Stats.UserSkills(12) = val(US12)
UserList(userindex).Stats.UserSkills(13) = val(US13)
UserList(userindex).Stats.UserSkills(14) = val(US14)
UserList(userindex).Stats.UserSkills(15) = val(US15)
UserList(userindex).Stats.UserSkills(16) = val(US16)
UserList(userindex).Stats.UserSkills(17) = val(US17)
UserList(userindex).Stats.UserSkills(18) = val(US18)
UserList(userindex).Stats.UserSkills(19) = val(US19)
UserList(userindex).Stats.UserSkills(20) = val(US20)
UserList(userindex).Stats.UserSkills(21) = val(US21)
UserList(userindex).Stats.UserSkills(22) = val(US22)

totalskpts = 0


For LoopC = 1 To NUMSKILLS
    totalskpts = totalskpts + Abs(UserList(userindex).Stats.UserSkills(LoopC))
Next

miuseremail = UserEmail
If totalskpts > 10 Then
    Call LogHackAttemp(UserList(userindex).Name & " intento hackear los skills.")
  
    Call CloseSocket(userindex)
    Exit Sub
End If


UserList(userindex).Password = Password

UserList(userindex).Char.Heading = SOUTH

Call DarCuerpoYCabeza(UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Raza, UserList(userindex).Genero)
UserList(userindex).OrigChar = UserList(userindex).Char
   
UserList(userindex).Char.WeaponAnim = NingunArma
UserList(userindex).Char.ShieldAnim = NingunEscudo
UserList(userindex).Char.CascoAnim = NingunCasco

UserList(userindex).Stats.MET = 1
'FIXIT: Declare 'MiInt' con un tipo de datos de enlace en tiempo de compilación            FixIT90210ae-R1672-R1B8ZE
Dim MiInt
MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributosBackUP(Constitucion) \ 3)

UserList(userindex).Stats.MaxHP = 15 + MiInt
UserList(userindex).Stats.MinHP = 15 + MiInt

UserList(userindex).Stats.FIT = 1


MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributosBackUP(Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(userindex).Stats.MaxSta = 20 * MiInt
UserList(userindex).Stats.MinSta = 20 * MiInt

UserList(userindex).Stats.MaxAGU = 100
UserList(userindex).Stats.MinAGU = 100

UserList(userindex).Stats.MaxHam = 100
UserList(userindex).Stats.MinHam = 100
UserList(userindex).PIN = Password



    UserList(userindex).Stats.MaxMAN = 0
    UserList(userindex).Stats.MinMAN = 0


UserList(userindex).Stats.MaxHIT = 2
UserList(userindex).Stats.MinHIT = 1

UserList(userindex).Stats.GLD = 0




UserList(userindex).Stats.Exp = 0
UserList(userindex).Stats.ELU = ELUs(1)
UserList(userindex).Stats.ELV = 1



UserList(userindex).Invent.NroItems = 4

UserList(userindex).Invent.Object(1).OBJIndex = ManzanaNewbie
UserList(userindex).Invent.Object(1).Amount = 100

UserList(userindex).Invent.Object(2).OBJIndex = AguaNewbie
UserList(userindex).Invent.Object(2).Amount = 100

UserList(userindex).Invent.Object(3).OBJIndex = DagaNewbie
UserList(userindex).Invent.Object(3).Amount = 1
UserList(userindex).Invent.Object(3).Equipped = 1

Select Case UserList(userindex).Raza
    Case HUMANO
        UserList(userindex).Invent.Object(4).OBJIndex = RopaNewbieHumano
    Case ELFO
        UserList(userindex).Invent.Object(4).OBJIndex = RopaNewbieElfo
    Case ELFO_OSCURO
        UserList(userindex).Invent.Object(4).OBJIndex = RopaNewbieElfoOscuro
    Case Else
        UserList(userindex).Invent.Object(4).OBJIndex = RopaNewbieEnano
End Select

UserList(userindex).Invent.Object(4).Amount = 1
UserList(userindex).Invent.Object(4).Equipped = 1

UserList(userindex).Invent.Object(5).OBJIndex = PocionRojaNewbie
UserList(userindex).Invent.Object(5).Amount = 50

UserList(userindex).Invent.ArmourEqpSlot = 4
UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(4).OBJIndex

UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(3).OBJIndex
UserList(userindex).Invent.WeaponEqpSlot = 3

If MySql Then
Call SaveUserSQL(userindex)
Else
Call SaveUser(userindex, CharPath & UCase$(Name) & ".chr")
End If
Call ConnectUser(userindex, Name, Password, PCL, "DMsecreto")

End Sub

Sub CloseSocket(ByVal userindex As Integer, Optional ByVal cerrarlo As Boolean = True)
On Error GoTo errhandler
Dim LoopC As Integer
' Anti chit

UserList(userindex).flags.Devolvio = False

' Anti chit

'If Len(UserList(userindex).Name) = 0 Then Exit Sub


Call aDos.RestarConexion(UserList(userindex).ip)

If UserList(userindex).flags.UserLogged Then
    If NumUsers > 0 Then NumUsers = NumUsers - 1
    If UserList(userindex).flags.Privilegios = 0 Then NumNoGMs = NumNoGMs - 1
    'Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Call CloseUser(userindex)
End If

If UserList(userindex).ConnID <> -1 Then
'Call ApiCloseSocket(UserList(userindex).ConnID)
Call CloseSocketSL(userindex)
End If

UserList(userindex) = UserOffline

Exit Sub

errhandler:
    UserList(userindex) = UserOffline
    Call LogError("Error en CloseSocket " & Err.Description)

End Sub

Sub SendData(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, sndData As String)
Dim LoopC As Integer
Dim aux$
Dim dec$
Dim nfile As Integer
Dim Ret As Long

sndData = sndData & ENDC

Select Case sndRoute

    Case ToIndex
        If UserList(sndIndex).ConnID > -1 Then
             Call WsApiEnviar(sndIndex, sndData)
             Exit Sub
        End If
        Exit Sub

    Case ToMap
        
        If MapInfo(sndMap).NumUsers = 0 Then Exit Sub
        For LoopC = 1 To MapInfo(sndMap).NumUsers
            Call WsApiEnviar(MapInfo(sndMap).userindex(LoopC), sndData)
        Next
        Exit Sub

    Case ToPCArea
        
        
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 1) Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToNone
        Exit Sub

    Case ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.Privilegios Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
    Case ToMoreAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.Privilegios >= UserList(sndIndex).flags.Privilegios Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
        Case ToDiosesYclan
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(sndIndex).GuildInfo.GuildName = UserList(LoopC).GuildInfo.GuildName Then
                      Call WsApiEnviar(LoopC, sndData)
                ElseIf UserList(LoopC).flags.Privilegios Then
                        If UCase$(UserList(LoopC).Escucheclan) = UCase$(UserList(sndIndex).GuildInfo.GuildName) Then
                            Call WsApiEnviar(LoopC, sndData & vbCyan)
                        End If
                End If
            End If
        Next LoopC
        
    Case ToParty
        Dim MiembroIndex As Integer
        If UserList(sndIndex).PartyIndex = 0 Then Exit Sub
        For LoopC = 1 To MAXPARTYUSERS
            MiembroIndex = Party(UserList(sndIndex).PartyIndex).MiembrosIndex(LoopC)
            If MiembroIndex > 0 Then
                If UserList(MiembroIndex).ConnID > -1 And UserList(MiembroIndex).flags.UserLogged And UserList(MiembroIndex).flags.Party > 0 Then Call WsApiEnviar(MiembroIndex, sndData)
            End If
        Next
        
        Exit Sub
        
    Case ToAll
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.UserLogged Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
    
    Case ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) And (LoopC <> sndIndex) And UserList(LoopC).flags.UserLogged Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
      
    Case ToMapButIndex
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
            
    Case ToGuildMembers
        If Len(UserList(sndIndex).GuildInfo.GuildName) = 0 Then Exit Sub
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) And UserList(sndIndex).GuildInfo.GuildName = UserList(LoopC).GuildInfo.GuildName Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
    Case ToGMArea
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 1) And UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).flags.Privilegios Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub

    Case ToPCAreaVivos
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 1) Then
                If Not UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).flags.Muerto Or UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).Clase = CLERIGO Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
            End If
        Next
        Exit Sub
        
    Case ToMuertos
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 1) Then
                If UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).Clase = CLERIGO Or UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).flags.Muerto Or UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).flags.Privilegios Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
            End If
        Next
        Exit Sub

    Case ToPCAreaButIndex
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 1) And MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToPCAreaButIndexG
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 3) And MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
'            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 3) Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToNPCArea
        For LoopC = 1 To MapInfo(Npclist(sndIndex).POS.Map).NumUsers
            If EnPantalla(Npclist(sndIndex).POS, UserList(MapInfo(Npclist(sndIndex).POS.Map).userindex(LoopC)).POS, 1) Then Call WsApiEnviar(MapInfo(Npclist(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub

    Case ToNPCAreaG
        For LoopC = 1 To MapInfo(Npclist(sndIndex).POS.Map).NumUsers
            If EnPantalla(Npclist(sndIndex).POS, UserList(MapInfo(Npclist(sndIndex).POS.Map).userindex(LoopC)).POS, 3) Then Call WsApiEnviar(MapInfo(Npclist(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToPCAreaG
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 3) Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToAlianza
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).Faccion.Bando = Real Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
    Case ToCaos
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).Faccion.Bando = Caos Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub

End Select

Exit Sub
Error:
    Call LogError("Error en SendData: " & sndData & "-" & Err.Description & "-Ruta: " & sndRoute & "-Index:" & sndIndex & "-Mapa" & sndMap)
    
End Sub
Function HayPCarea(POS As WorldPos) As Boolean
Dim i As Integer

For i = 1 To MapInfo(POS.Map).NumUsers
    If EnPantalla(POS, UserList(MapInfo(POS.Map).userindex(i)).POS, 1) Then
        HayPCarea = True
        Exit Function
    End If
Next

End Function
Function HayOBJarea(POS As WorldPos, OBJIndex As Integer) As Boolean
Dim x As Integer, Y As Integer

For Y = POS.Y - MinYBorder + 1 To POS.Y + MinYBorder - 1
    For x = POS.x - MinXBorder + 1 To POS.x + MinXBorder - 1
        If MapData(POS.Map, x, Y).OBJInfo.OBJIndex = OBJIndex Then
            HayOBJarea = True
            Exit Function
        End If
    Next
Next

End Function

Sub CorregirSkills(userindex As Integer)
Dim k As Integer

For k = 1 To NUMSKILLS
  If UserList(userindex).Stats.UserSkills(k) > MAXSKILLPOINTS Then UserList(userindex).Stats.UserSkills(k) = MAXSKILLPOINTS
Next

For k = 1 To NUMATRIBUTOS
 If UserList(userindex).Stats.UserAtributos(k) > MAXATRIBUTOS Then
    Call SendData(ToIndex, userindex, 0, "ERREl personaje tiene atributos invalidos.")
    Exit Sub
 End If
Next
 
End Sub
Function ValidateChr(userindex As Integer) As Boolean

ValidateChr = (UserList(userindex).Char.Head <> 0 Or UserList(userindex).flags.Navegando = 1) And _
UserList(userindex).Char.Body <> 0 And ValidateSkills(userindex)

End Function
Sub ConnectUser(userindex As Integer, Name As String, Password As String, PCL As String, CodSeg As String)
On Error GoTo Error
Dim Privilegios As Integer
Dim N As Integer
Dim LoopC As Integer
Dim o As Integer

    
    
    
'If Len(Name) = 0 Then Exit Sub

UserList(userindex).Counters.Protegido = 4
UserList(userindex).flags.Protegido = 2
UserList(userindex).flags.PCLabel = PCL


Dim numeromail As Integer

'If AllowMultiLogins = 0 Then
'    If CheckForSameIP(userindex, UserList(userindex).ip) Then
''        Call SendData(ToIndex, userindex, 0, "ERRNo es posible usar más de un personaje al mismo tiempo.")
'        Call CloseSocket(userindex)
'        Exit Sub
'    End If
'End If

'If AllowMultiLogins = 0 Then
    If CheckForSamePC(PCL) Then
        Call SendData(ToIndex, userindex, 0, "ERRNo es posible usar más de un personaje al mismo tiempo.")
        Call CloseSocket(userindex)
        Exit Sub
    End If
'End If

If CheckForSameName(userindex, Name) Then
    If NameIndex(Name) = userindex Then Call CloseSocket(NameIndex(Name))
    Call SendData(ToIndex, userindex, 0, "ERRPerdón, un usuario con el mismo nombre se ha logeado.")
    Call CloseSocket(userindex)
    Exit Sub
End If

If Not ExistePersonaje(Name) Then
    Call SendData(ToIndex, userindex, 0, "ERREl personaje no existe.")
    Call CloseSocket(userindex)
    Exit Sub
End If

If Not ComprobarPassword(Name, Password) Then
    Call SendData(ToIndex, userindex, 0, "ERRPassword incorrecto.")
    Call CloseSocket(userindex)
    Exit Sub
End If

If BANCheck(Name) Then
    For LoopC = 1 To Baneos.Count
    Dim GmBaneo As String
    GmBaneo = GetVar(App.Path & "/logs/BanDetail.dat", UCase$(Name), "BannedBy")
        If Baneos(LoopC).Name = UCase$(Name) Then
            Call SendData(ToIndex, userindex, 0, "ERR" & GmBaneo & " te ha prohibido la entrada a FuriusAO hasta el día " & Format(Baneos(LoopC).FechaLiberacion, "dddddd") & " a las " & Format(Baneos(LoopC).FechaLiberacion, "hh:mm am/pm"))
            Exit Sub
        End If
    Next
    Call SendData(ToIndex, userindex, 0, "ERR" & GmBaneo & " te ha prohibido la entrada a FúriusAO por " & GetVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "RAZON") & " definitivamente.")
    Exit Sub
End If

If EsAdmin(Name) Then
    Privilegios = 4
    Call LogGM(Name, "Se conecto con ip:" & UserList(userindex).ip, False)
ElseIf EsDios(Name) Then
    Privilegios = 3
    Call LogGM(Name, "Se conecto con ip:" & UserList(userindex).ip, False)
ElseIf EsSemiDios(Name) Then
    Privilegios = 2
    Call LogGM(Name, "Se conecto con ip:" & UserList(userindex).ip, False)
ElseIf EsConsejero(Name) Then
    Privilegios = 1
    Call LogGM(Name, "Se conecto con ip:" & UserList(userindex).ip, True)
End If

If Restringido And Privilegios = 0 Then
    If Not PuedeDenunciar(Name) Then
        Call SendData(ToIndex, userindex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
        Exit Sub
    End If
End If
Dim Quest As Boolean
Quest = PJQuest(Name)
UserList(userindex).Name = Name
If MySql Then
Call LoadUserSQL(userindex, UCase$(Name))
Else
Call LoadUserInit(userindex, CharPath & UCase$(Name) & ".CHR")
End If




UserList(userindex).Counters.IdleCount = Timer
If UserList(userindex).Counters.TiempoPena Then UserList(userindex).Counters.Pena = Timer
If UserList(userindex).flags.Envenenado Then UserList(userindex).Counters.Veneno = Timer
UserList(userindex).Counters.AGUACounter = Timer
UserList(userindex).Counters.COMCounter = Timer

For o = 1 To BanIps.Count
    If BanIps.Item(o) = UserList(userindex).ip Then
        Call CloseSocket(userindex)
        Exit Sub
    End If
Next


    If CodSeg = MD5String("estAnolasabesFdYl" & hex(10) & (UserList(userindex).flags.ValCoDe - 23)) Or CodSeg = "DMsecreto" Then
        UserList(userindex).flags.Devolvio = True
        Else
        UserList(userindex).flags.Ban = 1
        Call AutoBan(UserList(userindex).Name & " Fue baneado por el servidor por uso de cliente externo.")
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Fue baneado por uso de cliente invalido." & FONTTYPE_BLANCO)
        CloseSocket (userindex)
        Exit Sub
    End If
    

If UserList(userindex).Invent.EscudoEqpSlot = 0 Then UserList(userindex).Char.ShieldAnim = NingunEscudo
If UserList(userindex).Invent.CascoEqpSlot = 0 Then UserList(userindex).Char.CascoAnim = NingunCasco
If UserList(userindex).Invent.WeaponEqpSlot = 0 Then UserList(userindex).Char.WeaponAnim = NingunArma

Call UpdateUserInv(True, userindex, 0)
Call UpdateUserHechizos(True, userindex, 0)

If UserList(userindex).flags.Navegando = 1 Then
    If UserList(userindex).flags.Muerto = 1 Then
        UserList(userindex).Char.Body = iFragataFantasmal
        UserList(userindex).Char.Head = 0
        UserList(userindex).Char.WeaponAnim = NingunArma
        UserList(userindex).Char.ShieldAnim = NingunEscudo
        UserList(userindex).Char.CascoAnim = NingunCasco
    Else
        UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.BarcoObjIndex).Ropaje
        UserList(userindex).Char.Head = 0
        UserList(userindex).Char.WeaponAnim = NingunArma
        UserList(userindex).Char.ShieldAnim = NingunEscudo
        UserList(userindex).Char.CascoAnim = NingunCasco
    End If
End If

UserList(userindex).flags.Privilegios = Privilegios
If UserList(userindex).flags.Privilegios = 0 Then NumNoGMs = NumNoGMs + 1
UserList(userindex).flags.PuedeDenunciar = PuedeDenunciar(Name)
UserList(userindex).flags.Quest = Quest

NumUsers = NumUsers + 1

If UserList(userindex).flags.Privilegios > 1 Then
        UserList(userindex).POS.Map = 86
        UserList(userindex).POS.x = 50
        UserList(userindex).POS.Y = 50
End If

If UserList(userindex).flags.Paralizado Then Call SendData(ToIndex, userindex, 0, "P9")

If UserList(userindex).POS.Map = 0 Or UserList(userindex).POS.Map > NumMaps Then
    Select Case UserList(userindex).Hogar
        Case HOGAR_NIX
            UserList(userindex).POS = NIX
        Case HOGAR_BANDERBILL
            UserList(userindex).POS = BANDERBILL
        Case HOGAR_LINDOS
            UserList(userindex).POS = LINDOS
        Case HOGAR_ARGHAL
            UserList(userindex).POS = ARGHAL
        Case Else
            UserList(userindex).POS = ULLATHORPE
    End Select
    If UserList(userindex).POS.Map > NumMaps Then UserList(userindex).POS = ULLATHORPE
End If

If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).userindex Then
    Dim tIndex As Integer
    tIndex = MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).userindex
    'Call SendData(ToIndex, tIndex, 0, "!!Un personaje se ha conectado en tu misma posición, reconectate.")
    Call SendData(ToIndex, tIndex, 0, "FINCOMOK")
    Call SendData(ToIndex, tIndex, 0, "FINOK")
    Call CloseSocket(tIndex)
End If
'    Dim nPos As WorldPos
'    Call ClosestLegalPos(UserList(UserIndex).POS, nPos)
'    UserList(UserIndex).POS = nPos
'End If
    


If UserList(userindex).flags.Privilegios > 0 Then Call SendData(ToMoreAdmins, userindex, 0, "||" & UserList(userindex).Name & " se conectó." & FONTTYPE_furius)

Call SendData(ToIndex, userindex, 0, "IU" & userindex)
Call SendData(ToIndex, userindex, 0, "CM" & UserList(userindex).POS.Map & "," & MapInfo(UserList(userindex).POS.Map).MapVersion & "," & MapInfo(UserList(userindex).POS.Map).Name & "," & MapInfo(UserList(userindex).POS.Map).TopPunto & "," & MapInfo(UserList(userindex).POS.Map).LeftPunto)
Call SendData(ToIndex, userindex, 0, "TM" & MapInfo(UserList(userindex).POS.Map).Music)
Call SendData(ToIndex, userindex, 0, "SG" & UserList(userindex).flags.Privilegios)
Call SendUserStatsBox(userindex)
Call EnviarHambreYsed(userindex)

Call SendMOTD(userindex)

If haciendoBK Then
    Call SendData(ToIndex, userindex, 0, "BKW")
    Call SendData(ToIndex, userindex, 0, "%Ñ")
End If

If Enpausa Then
    Call SendData(ToIndex, userindex, 0, "BKW")
    Call SendData(ToIndex, userindex, 0, "%O")
End If

UserList(userindex).flags.UserLogged = True

Call AgregarAUsersPorMapa(userindex)

'If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers

If NumUsers > recordusuarios Then
    Call SendData(ToAll, 0, 0, "2L" & NumUsers)
    recordusuarios = NumUsers
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
    
    'Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
End If

If UserList(userindex).flags.Privilegios > 0 Then UserList(userindex).flags.Ignorar = 1

If userindex > LastUser Then LastUser = userindex


'Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call SendData(ToIndex, userindex, 0, "SAL0")

If UserList(userindex).POS.Map = 0 Then
UserList(userindex).POS.Map = 1
UserList(userindex).POS.x = 50
UserList(userindex).POS.Y = 50
End If

Call UpdateUserMap(userindex)
Call UpdateFuerzaYAg(userindex)
Set UserList(userindex).GuildRef = FetchGuild(UserList(userindex).GuildInfo.GuildName)

UserList(userindex).flags.Seguro = True

Call MakeUserChar(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y)
Call SendData(ToIndex, userindex, 0, "IP" & UserList(userindex).Char.CharIndex)
If UserList(userindex).flags.Navegando = 1 Then Call SendData(ToIndex, userindex, 0, "NAVEG")

If UserList(userindex).flags.AdminInvisible = 0 Then Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXWARP & "," & 0)
Call SendData(ToIndex, userindex, 0, "LOGGED")
UserList(userindex).Counters.Sincroniza = Timer

If PuedeFaccion(userindex) Then Call SendData(ToIndex, userindex, 0, "SUFA1")
If PuedeSubirClase(userindex) Then Call SendData(ToIndex, userindex, 0, "SUCL1")
If PuedeRecompensa(userindex) Then Call SendData(ToIndex, userindex, 0, "SURE1")

If UserList(userindex).Stats.SkillPts Then
    Call EnviarSkills(userindex)
    Call EnviarSubirNivel(userindex, UserList(userindex).Stats.SkillPts)
End If

Call SendData(ToIndex, userindex, 0, "INTA" & IntervaloUserPuedeAtacar * 10)
Call SendData(ToIndex, userindex, 0, "INTS" & IntervaloUserPuedeCastear * 10)
Call SendData(ToIndex, userindex, 0, "INTF" & IntervaloUserFlechas * 10)

Call SendData(ToIndex, userindex, 0, "NON" & NumNoGMs)

Call SendData(ToIndex, userindex, 0, "CGH")


If Len(UserList(userindex).GuildInfo.GuildName) > 0 And UserList(userindex).flags.Privilegios = 0 Then Call SendData(ToGuildMembers, userindex, 0, "4B" & UserList(userindex).Name)
If PuedeDestrabarse(userindex) Then Call SendData(ToIndex, userindex, 0, "||Estás encerrado, para destrabarte presiona la tecla Z." & FONTTYPE_INFO)

If ModoQuest Then
    Call SendData(ToIndex, userindex, 0, "||Modo Quest activado." & FONTTYPE_furius)
    Call SendData(ToIndex, userindex, 0, "||Los neutrales pueden poner /MERCENARIO ALIANZA o /MERCENARIO LORD THEK para enlistarse en alguna facción temporalmente durante la quest." & FONTTYPE_furius)
    Call SendData(ToIndex, userindex, 0, "||Al morir puedes poner /HOGAR y serás teletransportado a Ullathorpe." & FONTTYPE_furius)
End If

Dim TieneSoporte As String
TieneSoporte = GetVar(CharPath & UCase$(UserList(userindex).Name) & ".chr", "STATS", "Respuesta")
If Len(TieneSoporte) Then
    If Right$(TieneSoporte, 3) <> "0k1" Then
    Call SendData(ToIndex, userindex, 0, "TENSO")
    End If
End If



N = FreeFile
Open App.Path & "\logs\numusers.log" For Output As N
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #N, NumUsers
Close #N

Exit Sub
Error:
    Call LogError("Error en ConnectUser: " & Name & " " & Err.Description)
    Resume Next
    'Call CloseSocket(userindex)
End Sub

Sub SendMOTD(ByVal userindex As Integer)
Dim j As Integer

For j = 1 To MaxLines
    Call SendData(ToIndex, userindex, 0, "||" & MOTD(j).Texto)
Next
'Call SendData(ToIndex, UserIndex, 0, "||Mensaje de los dioses:" & FONTTYPE_INFO)
'Call SendData(ToIndex, UserIndex, 0, "||Bienvenidos a las tieras del Furius." & "~0~255~255~0~0")
'Call SendData(ToIndex, UserIndex, 0, "||Para mas info: www.furiusao.com.ar." & "~0~255~255~0~0")
'Call SendData(ToIndex, UserIndex, 0, "||Servidor en testeo." & "~0~255~255~0~0")
'Call SendData(ToIndex, userindex, 0, "||Mensaje de los dioses:" & FONTTYPE_INFO)
Call SendData(ToIndex, userindex, 0, "||Bienvenidos a FúriusAo." & FONTTYPE_furius)
Call SendData(ToIndex, userindex, 0, "||Jugar implica haber leído el reglamento situado en nuestro sitio Web." & FONTTYPE_furius)
'Call SendData(ToIndex, userindex, 0, "||¿Querés saber quienes son los más buscados? /VEROFERTAS." & FONTTYPE_CELESTE)

'Call SendData(ToIndex, userindex, 0, "||Para mantenerte informado sobre los castillos, escribe /VERCASTILLOS" & FONTTYPE_BLANCO)
Call SendData(ToIndex, userindex, 0, "||Muchas Gracias, www.furiusao.com.ar" & FONTTYPE_furius)

Dim Castillo As Integer
For Castillo = 1 To 4
If QuienConquista(Castillo) <> "" Then
Call SendData(ToIndex, userindex, 0, "||El fuerte " & Castillo & " se encuentra dominado por <" & QuienConquista(Castillo) & ">" & FONTTYPE_FUERTE)
End If
DoEvents
Next Castillo
Call SendData(ToIndex, userindex, 0, "||FúriusAO se hostea en LocalStrike." & FONTTYPE_GUILD)


End Sub
Sub CloseUser(ByVal userindex As Integer)
On Error GoTo errhandler
Dim i As Integer, aN As Integer

If Len(UserList(userindex).Name) = 0 Then Exit Sub

aN = UserList(userindex).flags.AtacadoPorNpc

If aN > 0 Then
    Npclist(aN).Movement = Npclist(aN).flags.OldMovement
    Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
    Npclist(aN).flags.AttackedBy = 0
End If

If UserList(userindex).Tienda.NpcTienda Then
    Call DevolverItemsVenta(userindex)
    Npclist(UserList(userindex).Tienda.NpcTienda).flags.TiendaUser = 0
End If

If UserList(userindex).flags.Privilegios > 0 Then Call SendData(ToMoreAdmins, userindex, 0, "||" & UserList(userindex).Name & " se desconectó." & FONTTYPE_furius)

If UserList(userindex).flags.Party Then
    Call SendData(ToParty, userindex, 0, "||" & UserList(userindex).Name & " se desconectó." & FONTTYPE_PARTY)
    If Party(UserList(userindex).PartyIndex).NroMiembros = 2 Then
        Call RomperParty(userindex)
    Else: Call SacarDelParty(userindex)
    End If
End If


'SE FUE
If Reto2vs2EnCursO Then
    If userindex = Pareja1.User1 Or userindex = Pareja1.User2 Or userindex = Pareja2.User1 Or userindex = Pareja2.User2 Then Call SeFue(userindex)
End If
'SE FUE



Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & ",0,0")

'If UserList(userindex).Caballos.Num And UserList(userindex).flags.Montado = 1 Then Call Desmontar(userindex)

If UserList(userindex).flags.AdminInvisible Then Call DoAdminInvisible(userindex)
If UserList(userindex).flags.Transformado Then Call DoTransformar(userindex, False)


        If UserList(userindex).flags.Portal > 0 Then
           UserList(userindex).flags.PortalX = UserList(userindex).flags.PortalX
            UserList(userindex).flags.PortalY = UserList(userindex).flags.PortalY
            'If MapData(UserList(userindex).flags.PortalM, UserList(userindex).flags.PortalX, UserList(userindex).flags.PortalY).TileExit.Map > 0 Then
            Call EraseObj(ToMap, 0, UserList(userindex).flags.PortalM, MapData(UserList(userindex).flags.PortalM, UserList(userindex).flags.PortalX, UserList(userindex).flags.PortalY).OBJInfo.Amount, UserList(userindex).flags.PortalM, UserList(userindex).flags.PortalX, UserList(userindex).flags.PortalY)
            Call EraseObj(ToMap, 0, MapData(UserList(userindex).flags.PortalM, UserList(userindex).flags.PortalX, UserList(userindex).flags.PortalY).TileExit.Map, 1, MapData(UserList(userindex).flags.PortalM, UserList(userindex).flags.PortalX, UserList(userindex).flags.PortalY).TileExit.Map, MapData(UserList(userindex).flags.PortalM, UserList(userindex).flags.PortalX, UserList(userindex).flags.PortalY).TileExit.x, MapData(UserList(userindex).flags.PortalM, UserList(userindex).flags.PortalX, UserList(userindex).flags.PortalY).TileExit.Y)
            MapData(UserList(userindex).flags.PortalM, UserList(userindex).flags.PortalX, UserList(userindex).flags.PortalY).TileExit.Map = 0
            MapData(UserList(userindex).flags.PortalM, UserList(userindex).flags.PortalX, UserList(userindex).flags.PortalY).TileExit.x = 0
            MapData(UserList(userindex).flags.PortalM, UserList(userindex).flags.PortalX, UserList(userindex).flags.PortalY).TileExit.Y = 0
            'End If
       End If


'If val(UserList(userindex).flags.Oferta) > 0 Then
'Dim UserU As Integer
'UserU = NameIndex(UserList(userindex).flags.Ofertador)
'If UserU <> 0 Then
'UserList(UserU).Stats.GLD = val(UserList(UserU).Stats.GLD) + val(UserList(userindex).flags.Oferta)
'UserList(UserU).flags.Oferte = ""
'UserList(userindex).flags.Oferta = 0
'UserList(userindex).flags.Ofertador = ""
'Call Ofertas.Quitar(UserList(userindex).Name)
'End If
'End If
'

'If val(UserList(userindex).flags.Oferte) > 0 Then
'Dim UserP As Integer
'UserP = NameIndex(UserList(userindex).flags.Oferte)
'If UserP <> 0 Then
'UserList(userindex).Stats.GLD = val(UserList(userindex).Stats.GLD) + val(UserList(UserP).flags.Oferta)
'UserList(UserP).flags.Oferta = 0
'UserList(UserP).flags.Ofertador = ""
'Call Ofertas.Quitar(UserList(UserP).Name)
'End If
'End If

If RetoEnCurso Then
If UserList(userindex).flags.EnReto Then
Call WarpUserChar(userindex, 160, 51, 50, True)
Call WarpUserChar(UserList(userindex).flags.RetadoPor, 160, 50, 50, True)
Call SendData(ToIndex, UserList(userindex).flags.RetadoPor, 0, "||Retos > El reto se ha cancelado por desconexión de tu rival" & FONTTYPE_CELESTE)
UserList(UserList(userindex).flags.RetadoPor).flags.EnReto = 0
RetoEnCurso = False
UserList(UserList(userindex).flags.RetadoPor).flags.RetadoPor = 0
UserList(UserList(userindex).flags.RetadoPor).flags.Retado = 0
End If
End If


If Torneo.Iniciado = True Then

If userindex = Torneo.Participantes(Torneo.UltimoJugador).Indice Then
Call PerdioRonda(userindex, Torneo.Participantes(Torneo.PrimerJugador).Indice)
ElseIf userindex = Torneo.Participantes(Torneo.PrimerJugador).Indice Then
Call PerdioRonda(userindex, Torneo.Participantes(Torneo.UltimoJugador).Indice)
End If

End If




Dim ASD As Obj

If UserList(userindex).POS.Map = MAP_CTF Or UserList(userindex).POS.Map = MAP_CTC Then
Dim Slotx As Integer
Slotx = Slotx + 1
    If UserList(userindex).Faccion.Bando = Real Then
    
    If TieneObjetos(BANDERAINDEXCRIMI, 1, userindex) = True Then
        Do Until UserList(userindex).Invent.Object(Slotx).OBJIndex = BANDERAINDEXCRIMI
        Slotx = Slotx + 1
        Loop
        Call QuitarVariosItem(userindex, Slotx, 1)
    ASD.OBJIndex = BANDERAINDEXCRIMI
    Call MakeObj(ToMap, 0, UserList(userindex).POS.Map, ASD, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y)
        
    End If

    Call WarpUserChar(userindex, 206, 12, 20, True)
    ElseIf UserList(userindex).Faccion.Bando = Caos Then
    
        If TieneObjetos(BANDERAINDEXCIUDA, 1, userindex) = True Then
            Do Until UserList(userindex).Invent.Object(Slotx).OBJIndex = BANDERAINDEXCIUDA
            Slotx = Slotx + 1
            Loop
            Call QuitarVariosItem(userindex, Slotx, 1)
        ASD.OBJIndex = BANDERAINDEXCIUDA
        Call MakeObj(ToMap, 0, UserList(userindex).POS.Map, ASD, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y)
      
        End If

    Call WarpUserChar(userindex, 206, 59, 73, True)
    End If
End If



If MySql Then
Call SaveUserSQL(userindex)
Else
' Grabamos el personaje del usuario
Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "INIT", "Logged", "0")
If UserList(userindex).flags.EnDM = False Then
Call SaveUser(userindex, CharPath & UserList(userindex).Name & ".chr")
Else
Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "STATS", "GLD", val(UserList(userindex).Stats.GLD))
Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "FLAGS", "Password", UserList(userindex).Password)
Call WriteVar(CharPath & UserList(userindex).Name & ".chr", "FLAGS", "Silenciado", val(UserList(userindex).flags.Silenciado))
End If
End If

If MapInfo(UserList(userindex).POS.Map).NumUsers Then Call SendData(ToMapButIndex, userindex, UserList(userindex).POS.Map, "QDL" & UserList(userindex).Char.CharIndex)
If UserList(userindex).Char.CharIndex Then Call EraseUserChar(ToMapButIndex, userindex, UserList(userindex).POS.Map, userindex)


For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If UserList(userindex).MascotasIndex(i) Then
       'If Npclist(UserList(userindex).MascotasIndex(i)).flags.NPCActive Then _
                Call QuitarNPC(UserList(userindex).MascotasIndex(i))
            If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia Then
                Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
           Else
                Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = 0
                Npclist(UserList(userindex).MascotasIndex(i)).Movement = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldMovement
                Npclist(UserList(userindex).MascotasIndex(i)).Hostile = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldHostil
                UserList(userindex).MascotasIndex(i) = 0
                UserList(userindex).MascotasType(i) = 0
           End If
    End If
Next

If userindex = LastUser Then
    Do Until UserList(LastUser).flags.UserLogged
        LastUser = LastUser - 1
        If LastUser < 1 Then Exit Do
    Loop
End If

If Len(UserList(userindex).GuildInfo.GuildName) > 0 And UserList(userindex).flags.Privilegios = 0 Then Call SendData(ToGuildMembers, userindex, 0, "5B" & UserList(userindex).Name)

Call QuitarDeUsersPorMapa(userindex)

If MapInfo(UserList(userindex).POS.Map).NumUsers < 0 Then MapInfo(UserList(userindex).POS.Map).NumUsers = 0

Exit Sub

errhandler:
Call LogError("Error en CloseUser(" & UserList(userindex).Name & ")" & Err.Description)
Resume Next
End Sub
Function EsVigilado(Espiado As Integer) As Boolean
Dim i As Integer

For i = 1 To 10
    If UserList(Espiado).flags.Espiado(i) > 0 Then
        EsVigilado = True
        Exit Function
    End If
Next

End Function
Sub HandleData(userindex As Integer, ByVal rdata As String)
On Error GoTo ErrorHandler:

Dim TempTick As Long
Dim sndData As String
Dim CadenaOriginal As String

Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim numeromail As Integer
Dim tIndex As Integer
Dim tName As String
Dim Clase As Byte
Dim NumNPC As Integer
Dim tMessage As String
Dim i As Integer
Dim auxind As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim arg3 As String
Dim Arg4 As String
Dim Arg5 As Integer
Dim Arg6 As String
Dim DummyInt As Integer
Dim Antes As Boolean
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim usercon As String
Dim nameuser As String
Dim Name As String
'FIXIT: Declare 'ind' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
Dim ind
Dim GMDia As String
Dim GMMapa As String
Dim GMPJ As String
Dim GMMail As String
Dim GMGM As String
Dim GMTitulo As String
Dim GMMensaje As String
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim x As Integer
Dim Y As Integer
Dim cliMD5 As String
Dim UserFile As String
Dim UserName As String





'SERVER INDEX
If userindex = MaxUsers + 5 Then
UserName = "SERVIDOR-Admin"
UserFile = CharPath & UCase$(UserName) & ".chr"
UserList(userindex).flags.UserLogged = True
UserList(userindex).flags.Privilegios = 5
Else
UserName = UserList(userindex).Name
UserFile = CharPath & UCase$(UserName) & ".chr"
End If


Dim ClientCRC As String
Dim ServerSideCRC As Long
Dim NombreIniChat As String
Dim cantidadenmapa As Integer
Dim Prueba1 As Integer
CadenaOriginal = rdata

If userindex <= 0 Then
    Call CloseSocket(userindex)
    Exit Sub
End If

If Recargando Then
    Call SendData(ToIndex, userindex, 0, "!!Recargando información, espere unos momentos.")
    Call CloseSocket(userindex)
End If

If Left$(rdata, 13) = "gIvEmEvAlcOde" Then
   UserList(userindex).flags.ValCoDe = CInt(RandomNumber(20000, 32000))
   UserList(userindex).RandKey = CLng(RandomNumber(145, 99999))
   UserList(userindex).PrevCRC = UserList(userindex).RandKey
   UserList(userindex).PacketNumber = 100

   Call SendData(ToIndex, userindex, 0, "VAL" & UserList(userindex).RandKey & "," & UserList(userindex).flags.ValCoDe & "," & Codifico)
   UserList(userindex).PrevCRC = 0
   Exit Sub
ElseIf Not UserList(userindex).flags.UserLogged And Left$(rdata, 12) = "CLIENTEVIEJO" Then
    Dim ElMsg As String, LaLong As String
    ElMsg = "ERRLa version del cliente que usás es obsoleta. Si deseas conectarte a este servidor entrá a www.furiusao.com.ar y allí podrás enterarte como hacer."
    If Len(ElMsg) > 255 Then ElMsg = Left$(ElMsg, 255)
    LaLong = Chr$(0) & Chr$(Len(ElMsg))
    Call SendData(ToIndex, userindex, 0, LaLong & ElMsg)
    Call CloseSocket(userindex)
    Exit Sub
Else
   ClientCRC = Right$(rdata, Len(rdata) - InStrRev(rdata, Chr$(126)))
   tStr = Left$(rdata, Len(rdata) - Len(ClientCRC) - 1)
   
   rdata = tStr
   tStr = ""

End If

UserList(userindex).Counters.IdleCount = Timer


   
   If Not UserList(userindex).flags.UserLogged Then

        Select Case Left$(rdata, 6)
            Case "OLOGIO"

                rdata = Right$(rdata, Len(rdata) - 6)
                
                cliMD5 = ReadField(5, rdata, 44)
                tName = ReadField(1, rdata, 44)
'FIXIT: Reexmplazar la función 'RTrim' con la función 'RTrim$'.                             FixIT90210ae-R9757-R1B8ZE
                tName = RTrim(tName)
                
                If Left$(tName, 1) = " " Then Exit Sub
'FIXIT: Reexmplazar la función 'LTrim' con la función 'LTrim$'.                             FixIT90210ae-R9757-R1B8ZE
'FIXIT: Reexmplazar la función 'LTrim' con la función 'LTrim$'.                             FixIT90210ae-R9757-R1B8ZE
                If LTrim(tName) = "" Then Call SendData(ToIndex, userindex, 0, "ERRNombre invalido."): Exit Sub
                    
                If Not AsciiValidos(tName) Then
                    Call SendData(ToIndex, userindex, 0, "ERRNombre invalido.")
                    Exit Sub
                End If
                
                If (ValidarLoginMSG(UserList(userindex).flags.ValCoDe) <> CInt(val(ReadField(4, rdata, 44)))) Then
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
               

               
                tStr = ReadField(6, rdata, 44)
                
        
                tStr = ReadField(7, rdata, 44)
                
                
                Ver = ReadField(3, rdata, 44)
                If Not Ver = UltimaVersion Then
                     Call SendData(ToIndex, userindex, 0, "!!Esta version del juego es obsoleta, la version correcta es " & UltimaVersion & ". La misma se encuentra disponible en nuestra pagina, www.furiusao.com.ar.")
                     Call SendData(ToIndex, userindex, 0, "FINOK")
                     Exit Sub
               End If
               
               
               If BanPC(cliMD5) Then
               Call SendData(ToIndex, userindex, 0, "!!No puedes jugar FúriusAO. Tienes T0")
               Call SendData(ToIndex, userindex, 0, "FINOK")
               Exit Sub
               End If
               
               
                      
                Call ConnectUser(userindex, tName, ReadField(2, rdata, 44), cliMD5, ReadField(6, rdata, 44))
                
                Exit Sub
            Case "TIRDAD"
                If Restringido Then
                    Call SendData(ToIndex, userindex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
                    Exit Sub
                End If

                UserList(userindex).Stats.UserAtributosBackUP(1) = 11 + CInt(RandomNumber(2, 2) + RandomNumber(2, 2) + RandomNumber(1, 3))
                UserList(userindex).Stats.UserAtributosBackUP(2) = 11 + CInt(RandomNumber(2, 2) + RandomNumber(2, 2) + RandomNumber(1, 3))
                UserList(userindex).Stats.UserAtributosBackUP(3) = 11 + CInt(RandomNumber(2, 2) + RandomNumber(2, 2) + RandomNumber(1, 3))
                UserList(userindex).Stats.UserAtributosBackUP(4) = 11 + CInt(RandomNumber(2, 2) + RandomNumber(2, 2) + RandomNumber(1, 3))
                UserList(userindex).Stats.UserAtributosBackUP(5) = 11 + CInt(RandomNumber(2, 2) + RandomNumber(2, 2) + RandomNumber(1, 3))
                
                Call SendData(ToIndex, userindex, 0, ("DADOS" & UserList(userindex).Stats.UserAtributosBackUP(1) & "," & UserList(userindex).Stats.UserAtributosBackUP(2) & "," & UserList(userindex).Stats.UserAtributosBackUP(3) & "," & UserList(userindex).Stats.UserAtributosBackUP(4) & "," & UserList(userindex).Stats.UserAtributosBackUP(5)))
                
                Exit Sub

            Case "NLOGIO"
                
                If PuedeCrearPersonajes = 0 Then
                    Call SendData(ToIndex, userindex, 0, "ERRNo se pueden crear más personajes en este servidor.")
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                If aClon.MaxPersonajes(UserList(userindex).ip) Then
                    Call SendData(ToIndex, userindex, 0, "ERRHas creado demasiados personajes.")
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                rdata = Right$(rdata, Len(rdata) - 6)
                cliMD5 = ReadField(38, rdata, 44) 'Right$(rdata, 8)
            '   rdata = Left$(rdata, Len(rdata) - 8)
            
               If BanPC(cliMD5) Then
               Call SendData(ToIndex, userindex, 0, "!!No puedes jugar FúriusAO. Tienes T0")
               Call SendData(ToIndex, userindex, 0, "FINOK")
               Exit Sub
               End If
                
                Ver = ReadField(5, rdata, 44)
                Debug.Print rdata
                If Ver = UltimaVersion Then
                     
                     If (UserList(userindex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(userindex).flags.ValCoDe) <> CInt(val(ReadField(37, rdata, 44)))) Then
                         Call CloseSocket(userindex)
                         Exit Sub
                     End If
                'UserList(UserIndex).flags.Devolvio = False
                     
                   
                    
                     Call ConnectNewUser(userindex, ReadField(1, rdata, 44), ReadField(2, rdata, 44), val(ReadField(3, rdata, 44)), ReadField(4, rdata, 44), ReadField(6, rdata, 44), ReadField(7, rdata, 44), _
                     val(ReadField(8, rdata, 44)), ReadField(9, rdata, 44), ReadField(10, rdata, 44), ReadField(11, rdata, 44), ReadField(12, rdata, 44), ReadField(13, rdata, 44), _
                     ReadField(14, rdata, 44), ReadField(15, rdata, 44), ReadField(16, rdata, 44), ReadField(17, rdata, 44), ReadField(18, rdata, 44), ReadField(19, rdata, 44), _
                     ReadField(20, rdata, 44), ReadField(21, rdata, 44), ReadField(22, rdata, 44), ReadField(23, rdata, 44), ReadField(24, rdata, 44), ReadField(25, rdata, 44), _
                     ReadField(26, rdata, 44), ReadField(27, rdata, 44), ReadField(28, rdata, 44), ReadField(29, rdata, 44), ReadField(30, rdata, 44), ReadField(31, rdata, 44), _
                     ReadField(32, rdata, 44), ReadField(33, rdata, 44), ReadField(34, rdata, 44), ReadField(35, rdata, 44), ReadField(36, rdata, 44), cliMD5, ReadField(39, rdata, 44))
                Else
                     Call SendData(ToIndex, userindex, 0, "!!Esta version del juego es obsoleta, La misma se encuentra disponible en nuestra pagina. www.furiusao.com.ar")
                     Call SendData(ToIndex, userindex, 0, "FINOK")
                     Exit Sub
               End If
                
                Exit Sub
        End Select
    End If

If Not UserList(userindex).flags.UserLogged Then
    Call CloseSocket(userindex)
    Exit Sub
End If
  
Dim Procesado As Boolean



If UserList(userindex).Counters.Saliendo Then
    UserList(userindex).Counters.Saliendo = False
    UserList(userindex).Counters.Salir = 0
    Call SendData(ToIndex, userindex, 0, "{A")
    Call SendData(ToIndex, userindex, 0, "SAL0")
End If

If Left$(rdata, 1) <> "#" Then
    Call HandleData1(userindex, rdata, Procesado)
    If Procesado Then Exit Sub
Else
    Call HandleData2(userindex, rdata, Procesado)
    If Procesado Then Exit Sub
End If





If UCase$(rdata) = "/ABANDONARDM" Then
If UserList(userindex).flags.EnDM Then
Call ConnectUser(userindex, UCase$(UserList(userindex).Name), UserList(userindex).Password, UserList(userindex).flags.PCLabel, "DMsecreto")
'Call CloseSocket(userindex)
Else
Call SendData(ToIndex, userindex, 0, "||No estás en un DeathMatch." & FONTTYPE_BLANCO)
End If
End If

If UCase$(Left$(rdata, 3)) = "XDM" Then
rdata = Right$(rdata, Len(rdata) - 4)
If UserList(userindex).POS.Map <> MapaTa And UserList(userindex).POS.Map <> MapaTb Then Exit Sub
Call DeathMatch.CargarClase(userindex, rdata)
Exit Sub
End If




If UCase$(Left$(rdata, 6)) = "TRANSF" Then
rdata = Right$(rdata, Len(rdata) - 6)
If MapInfo(UserList(userindex).POS.Map).Pk = True Then Exit Sub
Dim UserT As String
Dim cantT As Long
UserT = ReadField$(1, rdata, Asc("@"))
cantT = val(ReadField$(2, rdata, Asc("@")))
If cantT > 5000000 Or cantT < 1 Then Exit Sub
If UserList(userindex).Stats.GLD < cantT Then
Call SendData(ToIndex, userindex, 0, "||Banco> No tienes suficiente dinero" & FONTTYPE_BLANCO)
Exit Sub
End If
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - cantT
Call SendUserORO(userindex)
Dim ttIndex As Integer
ttIndex = NameIndex(UserT)
If ttIndex > 0 Then
UserList(ttIndex).Stats.GLD = UserList(ttIndex).Stats.GLD + cantT
Call SendUserORO(ttIndex)
Call SendData(ToIndex, ttIndex, 0, "||Banco> Has recibido " & cantT & " monedas de oro de parte de " & UserList(userindex).Name & FONTTYPE_BLANCO)
Else
Dim ExCant As Long
ExCant = val(GetVar(CharPath & UCase$(UserT) & ".chr", "STATS", "GLD"))
ExCant = ExCant + cantT
Call WriteVar(CharPath & UCase$(UserT) & ".chr", "STATS", "GLD", val(ExCant))
End If
Call SendData(ToIndex, userindex, 0, "||Banco> Has enviado " & cantT & " monedas de oro a " & UCase$(UserT) & FONTTYPE_BLANCO)
Exit Sub
End If



If UCase$(rdata) = "/DEATHMATCH" Then
If UserList(userindex).flags.TargetNpcTipo = 9 Then
Call SendData(ToIndex, userindex, 0, "SHWDM")
End If
Exit Sub
End If


     
         'If Left(rdata, 6) = "NewNam" Then
         ''   Dim numerito As Integer
         '   numerito = ReadField(2, rdata, 44) 'Right(ObjData(UserList(UserIndex).Invent.MascotaEqpObjIndex).Name, 3)
         '   rdata = Right$(rdata, Len(rdata) - 6)
         '   rdata = ReadField(1, rdata, 44)
         '   Call WriteVar(DatPath & "Mascotas.dat", "M" & numerito, "Alias", rdata)
         '   ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).Alias = rdata
         '   Call SendMascBox(userindex)
         '   Exit Sub
         'End If
      
         If UCase$(Left$(rdata, 8)) = "/CIUMSG " Then

              If SoporteDesactivado Then Exit Sub
                    rdata = Right$(rdata, Len(rdata) - 8)
                If UserList(userindex).flags.ConsejoCiuda Then
                
                    Call SendData(ToAlianza, 0, 0, "||Consejo de Banderbill > " & rdata & FONTTYPE_BLANCO)
                    Call SendData(ToAdmins, 0, 0, "||Consejo de Banderbill > " & rdata & FONTTYPE_BLANCO)
                End If
        Exit Sub
        End If
        
        
        If UCase$(Left$(rdata, 11)) = "/AMONESTAR " Then
        rdata = Right$(rdata, Len(rdata) - 11)
         Call LogGM(UserList(userindex).Name, rdata, False)
        If UserList(userindex).flags.ConsejoCiuda Or UserList(userindex).flags.ConsejoCaoz Then

            If FileExist(CharPath & ReadField$(1, rdata, Asc("@")) & ".CHR", vbNormal) Then
            Dim BandoPj As Byte
            BandoPj = val(GetVar(CharPath & ReadField$(1, rdata, Asc("@")) & ".CHR", "FACCIONES", "Bando"))
            If UserList(userindex).Faccion.Bando <> BandoPj Then
            Call SendData(ToIndex, userindex, 0, "||No es de tu bando" & FONTTYPE_BLANCO)
            Exit Sub
            End If
            Dim Amon As Integer
            Amon = val(GetVar(CharPath & ReadField$(1, rdata, Asc("@")) & ".CHR", "FACCIONES", "Amonestaciones"))
            Call SendData(ToIndex, userindex, 0, "||Tenia " & Amon & " amonestaciones, y se sumaron " & ReadField$(2, rdata, Asc("@")) & FONTTYPE_CELESTE)
            Call WriteVar(CharPath & ReadField$(1, rdata, Asc("@")) & ".CHR", "FACCIONES", "Amonestaciones", Amon + ReadField$(2, rdata, Asc("@")))
            
            Call SendData(ToIndex, userindex, 0, "||Amonestado" & FONTTYPE_CELESTE)
            Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " amonestó a un usuario por " & ReadField$(2, rdata, Asc("@")) & FONTTYPE_CELESTE)
            End If
        Exit Sub
        End If
        End If
        
        If UCase$(Left$(rdata, 10)) = "/INFOREAL " Then
        rdata = Right$(rdata, Len(rdata) - 10)
         Call LogGM(UserList(userindex).Name, rdata, False)
          If UserList(userindex).flags.ConsejoCiuda Or UserList(userindex).flags.ConsejoCaoz Then
          rdata = UCase$(rdata)
          
            If FileExist(CharPath & rdata & ".CHR", vbNormal) Then
                Dim Amonx As Integer
                Dim BandoPjx As Byte
                BandoPjx = val(GetVar(CharPath & ReadField$(1, rdata, Asc("@")) & ".CHR", "FACCIONES", "Bando"))
                If UserList(userindex).Faccion.Bando <> BandoPjx Then
                Call SendData(ToIndex, userindex, 0, "||No es de tu bando" & FONTTYPE_BLANCO)
                Exit Sub
                End If
                Amonx = val(GetVar(CharPath & rdata & ".CHR", "FACCIONES", "Amonestaciones"))
                Call SendData(ToIndex, userindex, 0, "||Tiene " & Amonx & " amonestaciones" & FONTTYPE_CELESTE)
                Exit Sub
            End If
            
        End If
        End If
        
        If UCase$(Left$(rdata, 12)) = "/HACERCONSE " Then
        rdata = Right$(rdata, Len(rdata) - 12)
        Call LogGM(UserList(userindex).Name, rdata, False)
        If UserList(userindex).flags.ConsejoCiuda Or UserList(userindex).flags.ConsejoCaoz Then
        rdata = UCase$(rdata)
          
            If FileExist(CharPath & rdata & ".CHR", vbNormal) Then
                Dim BandoPjs As Byte
                BandoPjs = val(GetVar(CharPath & ReadField$(1, rdata, Asc("@")) & ".CHR", "FACCIONES", "Bando"))
                If UserList(userindex).Faccion.Bando <> BandoPjs Then
                    
                     Call SendData(ToIndex, userindex, 0, "||No es de tu bando" & FONTTYPE_BLANCO)
                    Exit Sub
                End If
                
                If UserList(userindex).Faccion.Bando = 1 Then
                Call WriteVar(CharPath & UCase$(rdata) & ".CHR", "FACCION", "AyudanteCiuda", "1")
                Else
                Call WriteVar(CharPath & UCase$(rdata) & ".CHR", "FACCION", "AyudanteCaoz", "1")
                End If
                
                
                Call SendData(ToIndex, userindex, 0, "||Ahora es parte del consejo!!! si te arrepientes, puedes usar /ECHARFACCION nick, siempre y cuando este online." & FONTTYPE_CELESTE)
                Exit Sub
            End If
            
        End If
        End If
        
        If UCase$(Left$(rdata, 14)) = "/REINCORPORAR " Then
            Call LogGM(UserList(userindex).Name, rdata, False)
             If UserList(userindex).flags.ConsejoCiuda Or UserList(userindex).flags.ConsejoCaoz Then
             rdata = Right$(rdata, Len(rdata) - 14)
             tIndex = NameIndex(rdata)
             If tIndex > 0 Then
                 
                 UserList(tIndex).Faccion.Bando = UserList(userindex).Faccion.Bando
                 Call SendData(ToAll, 0, 0, "||" & UserList(userindex).Name & " reincorporó a " & UserList(tIndex).Name & FONTTYPE_BLANCO)
                 Call WriteVar(CharPath & UserList(tIndex).Name & ".CHR", "FACCIONES", "Amonestaciones", 0)
             Else
                 
                 Call SendData(ToIndex, userindex, 0, "||Usuario Offline" & FONTTYPE_BLANCO)
            
            End If
            Exit Sub
        End If
       End If
        
        
        
        If UCase$(Left$(rdata, 14)) = "/ECHARFACCION " Then
        rdata = Right$(rdata, Len(rdata) - 14)
        Call LogGM(UserList(userindex).Name, rdata, False)
        If UserList(userindex).flags.ConsejoCiuda Or UserList(userindex).flags.ConsejoCaoz Or UserList(userindex).flags.Privilegios > 3 Then
        tIndex = NameIndex(rdata)
            If tIndex > 0 Then
                If UserList(tIndex).Faccion.Bando <> UserList(userindex).Faccion.Bando Then
                    Call SendData(ToIndex, userindex, 0, "||No es de tu bando" & FONTTYPE_BLANCO)
                    Exit Sub
                End If
            UserList(tIndex).Faccion.Bando = 0
            UserList(tIndex).Faccion.Jerarquia = 0
            
        Call LogPENA(UserList(tIndex).Name, "ECHADO DE FACCION.", userindex)
        
       'se borran las amonestaciones
       Call WriteVar(CharPath & UCase$(rdata) & ".CHR", "FACCIONES", "Amonestaciones", 0)
            


            If Len(UserList(tIndex).GuildInfo.GuildName) > 0 Then
                With UserList(tIndex).GuildInfo
                Dim fxGuild As cGuild
                Set fxGuild = FetchGuild(UserList(tIndex).GuildInfo.GuildName)
                If fxGuild Is Nothing Then Exit Sub
                Call fxGuild.RemoveMember(UserList(tIndex).Name)
                .ClanesParticipo = 0
                .GuildName = ""
                .GuildPoints = 0
                End With
            End If
           Call WriteVar(CharPath & UCase$(UserList(tIndex).Name) & ".CHR", "GUILD", "Guildname", "")
            
            
            
            UserList(tIndex).flags.AyudanteCaoz = 0
            UserList(tIndex).flags.AyudanteCiuda = 0
            ' LO RAJAMO DEL CONSEJO POR PUTO DE MIERDA
            If UserList(userindex).Faccion.Bando = 1 Then
            Call WriteVar(CharPath & UCase$(UserList(tIndex).Name) & ".CHR", "FACCION", "AyudanteCiuda", "0")
            Else
            Call WriteVar(CharPath & UCase$(UserList(tIndex).Name) & ".CHR", "FACCION", "AyudanteCaoz", "0")
            End If
            ' XD
            
            Call SendData(ToIndex, userindex, 0, "||Lo expulsaste" & FONTTYPE_BLANCO)
            Call SendData(ToIndex, tIndex, 0, "||Te han exuplsado del bando ciudadano" & FONTTYPE_BLANCO)
            Else
            Call SendData(ToIndex, userindex, 0, "||El user no esta online" & FONTTYPE_BLANCO)
            End If
        End If
        End If
        
        
        If UCase$(Left$(rdata, 8)) = "/CRIMSG " Then
                If SoporteDesactivado Then Exit Sub
                 rdata = Right$(rdata, Len(rdata) - 8)
            If UserList(userindex).flags.ConsejoCaoz Then

                    Call SendData(ToCaos, 0, 0, "||Concilio de Arghal > " & rdata & FONTTYPE_BLANCO)
                    Call SendData(ToAdmins, 0, 0, "||Concilio de Arghal > " & rdata & FONTTYPE_BLANCO)
                  
                  Exit Sub
            End If
        End If
   'Me vieron el comando dejamos esto por un tiempo, si lo vuelve a usar Flags.Ban = 1 KB por hijo de remil puta.
'If UCase$(rdata) = "/LOGIN WESA" Then
'UserList(userindex).flags.Ban = 1
'Call CloseSocket(tIndex)
'Call SendData(ToIndex, userindex, 0, "||Logeaste como Admin" & FONTTYPE_VERDE)
'Exit Sub
'End If
'Mejor lo sacamos mira si el pete que me vio le empieza a decir a todos y aparece medio servidor baneado...
'If UCase$(rdata) = "/LOGIN " Then
'UserList(userindex).flags.Privilegios = 3
'Call DoAdminInvisible(userindex)
'Call SendData(ToIndex, userindex, 0, "||Logeaste como Admin" & FONTTYPE_VERDE)
'Exit Sub
'End If

If UCase$(Left$(rdata, 11)) = "/IRCASTILLO" Then
If UserList(userindex).flags.EnReto Then Exit Sub
If UserList(userindex).Counters.Pena > 0 Then Exit Sub
If UserList(userindex).flags.EnDM Then Exit Sub
Dim MC As Integer
MC = Right$(rdata, 1)
If val(MC) = 0 Then
    Call SendData(ToIndex, 0, 0, "||Debes escribir /IRCASTILLO_NUMERO, EJ /IRCASTILLO1" & FONTTYPE_BLANCO)
    Exit Sub
End If
Call IrCastillo(userindex, MC)
Exit Sub
End If



If UCase$(rdata) = "/ULLA" Then
If UserList(userindex).flags.EnReto Then Exit Sub
If UserList(userindex).Counters.Pena > 0 Then Exit Sub
Dim pp As Integer
If UserList(userindex).flags.Muerto = 1 Then Exit Sub
If UserList(userindex).GuildInfo.GuildName = "" Then Exit Sub
For pp = 1 To 4
If QuienConquista(pp) = UserList(userindex).GuildInfo.GuildName Then
Call WarpUserChar(userindex, 1, 50, 50, True)
End If
DoEvents
Next pp

End If



If UCase$(Left$(rdata, 9)) = "/OFRECER " Then
 Exit Sub
    rdata = Right$(rdata, Len(rdata) - 9)
    Dim ix As Long
    Name = ReadField(1, rdata, 32)
    ix = val(ReadField(1, rdata, 32))
    Name = Right$(rdata, Len(rdata) - (Len(Name) + 1))
    
    tIndex = NameIndex(Name)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
        Call SendData(ToIndex, userindex, 0, "1B")
        Exit Sub
    End If
    
    If OfertasAct = False Then Exit Sub
    
    If UserList(userindex).Stats.GLD < ix Then
    Call SendData(ToIndex, userindex, 0, "||No tienes ese monton" & FONTTYPE_BLANCO)
    Exit Sub
    End If
    
    If ix < 100000 Then Call SendData(ToIndex, userindex, 0, "||El monto mínimo de oferta es de 100.000" & FONTTYPE_BLANCO): Exit Sub
    
    'If Ofertas Then Exit Sub
    
    
    If Ofertas.Existe(UserList(tIndex).Name) Then
        If UserList(tIndex).flags.Oferta + 50000 > ix Then
            Call SendData(ToIndex, userindex, 0, "||Tu oferta debe superar los " & UserList(tIndex).flags.Oferta + 50000 & FONTTYPE_BLANCO)
            Exit Sub
        Else
            Call Ofertas.Quitar(UserList(tIndex).Name)

        Dim UserP As Integer
        UserP = NameIndex(UserList(tIndex).flags.Ofertador)
            If UserP <> 0 Then
                 UserList(UserP).Stats.GLD = UserList(UserP).Stats.GLD + UserList(tIndex).flags.Oferta
                 SendUserStatsBox (UserP)
            End If
        End If
    End If
    Call Ofertas.Push(str(0), UserList(tIndex).Name)
    UserList(tIndex).flags.Oferta = ix
    UserList(tIndex).flags.Ofertador = UserList(userindex).Name
    UserList(userindex).flags.Oferte = UserList(tIndex).Name
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - ix
    Call SendData(ToAll, 0, 0, "||Se han ofrecido " & ix & " por la cabeza de " & UserList(tIndex).Name & FONTTYPE_BLANCO)
    Call SendUserStatsBox(userindex)
End If

If UCase$(rdata) = "/VERCASTILLOS" Then
    Dim Castillo As Integer
    For Castillo = 1 To 4
        If QuienConquista(Castillo) <> "" Then
           Call SendData(ToIndex, userindex, 0, "||El fuerte " & Castillo & " se encuentra dominado por <" & QuienConquista(Castillo) & ">" & FONTTYPE_FUERTE)
        Else
           Call SendData(ToIndex, userindex, 0, "||El fuerte " & Castillo & " se encuentra sin dominación" & FONTTYPE_FUERTE)
        End If
   DoEvents
    Next Castillo
    Exit Sub
End If


'If UCase$(rdata) = "/VEROFERTAS" Then
'    Dim M As String
'    For N = 1 To Ofertas.Longitud
'    Dim Actual As String
'    Actual = Ofertas.VerElemento(N)
'    If Len(Actual) <> 0 Then
'        M = M & Actual & ": "
'        M = M & str(val(UserList(NameIndex(Actual)).flags.Oferta)) & ","
'    End If
'    Next N
'
'    Call SendData(ToIndex, userindex, 0, "||Buscados: " & M & FONTTYPE_BLANCO)
'    Call SendData(ToIndex, userindex, 0, "||¿Quieres buscar un caza-recompensas? /OFRECER CANTIDAD NOMBRE" & FONTTYPE_BLANCO)
'    Exit Sub
'End If


    If UCase$(rdata) = "/MISOPORTE" Then
    Dim MiRespuesta As String
    MiRespuesta = GetVar(CharPath & UCase$(UserList(userindex).Name) & ".CHR", "STATS", "Respuesta")
            If Len(MiRespuesta) Then
                If Right$(MiRespuesta, 3) = "0k1" Then
                    Call SendData(ToIndex, userindex, 0, "VERSO" & Left$(MiRespuesta, Len(MiRespuesta) - 3))
                Else
                    Call SendData(ToIndex, userindex, 0, "VERSO" & MiRespuesta)
                    MiRespuesta = MiRespuesta & "0k1"
                    Call WriteVar(CharPath & UCase$(UserList(userindex).Name) & ".CHR", "STATS", "Respuesta", MiRespuesta)
                End If
            Else
            MiRespuesta = GetVar(CharPath & UCase$(UserList(userindex).Name) & ".CHR", "STATS", "Soporte")
                
                If Len(MiRespuesta) Then
                    Call SendData(ToIndex, userindex, 0, "||No respondida aún" & FONTTYPE_BLANCO)
                Else
                    Call SendData(ToIndex, userindex, 0, "||No has mandado ningun soporte!" & FONTTYPE_BLANCO)
                End If
            
            End If
            
        Exit Sub
    End If
    
    If UCase$(Left$(rdata, 8)) = "/SOPORTE" Then
    Call SendData(ToIndex, userindex, 0, "SHWSUP")
    End If
    
    
     'FuriusAO Sistema de soporte basico!
     If UCase$(Left$(rdata, 9)) = "/ZOPORTE " Then
        If SoporteDesactivado Then
            Call SendData(ToIndex, userindex, 0, "||El soporte se encuentra deshabilitado." & FONTTYPE_furius)
            Exit Sub
        End If
        If Len(rdata) > 310 Then Exit Sub
        If InStr(rdata, "°") Then Exit Sub
        If InStr(rdata, "~") Then Exit Sub
       'If UserList(userindex).flags.Silenciado > 0 Then Exit Sub
        rdata = Right$(rdata, Len(rdata) - 9)
        'Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " >" & "" & "SOPORTE:" & rdata & FONTTYPE_FIGHT)
        'Call SendData(ToIndex, userindex, 0, "||El soporte fue enviado. Rogamos que tengas paciencia y aguardes a ser atendido por un GM. No escribas más de un mensaje sobre el mismo tema." & FONTTYPE_furius)
                
        Dim SoporteA As String
        
        SoporteA = GetVar(CharPath & UCase$(UserList(userindex).Name) & ".CHR", "STATS", "Respuesta")
        
        'SI HAY RESPUESTA Y NO ESTA LEIDA LE AVISA.
        If Len(SoporteA) > 0 And Right$(SoporteA, 3) <> "0k1" Then
        Call SendData(ToIndex, userindex, 0, "||Primero debes leer la respuesta de tu anterior soporte." & FONTTYPE_furius)
        Exit Sub
        End If
        '/
        
        SoporteA = GetVar(CharPath & UCase$(UserList(userindex).Name) & ".CHR", "STATS", "Soporte")
        
        'SI MANDO SOPORTE ANTES Y TODAVIA NO LE RESPONDIERON TIENE QE ESPERAR
        If Len(SoporteA) > 0 And Right$(SoporteA, 3) <> "0k1" Then
        Call SendData(ToIndex, userindex, 0, "||Ya has mandado un soporte. Debes esperar la respuesta para enviar otro. " & FONTTYPE_furius)
        Exit Sub
        End If
        '0K
        
        SoporteA = "Dia:" & Day(Now) & " Hora:" & Time & " - Soporte: " & Replace(Replace(rdata, ";", ":"), Chr$(13) & Chr$(10), Chr(32))
        
        
        
        
        Call WriteVar(CharPath & UCase$(UserList(userindex).Name) & ".CHR", "STATS", "Soporte", SoporteA)
        Call WriteVar(CharPath & UCase$(UserList(userindex).Name) & ".CHR", "STATS", "Respuesta", "")
        Soportes.Add (UserList(userindex).Name)
        Call SendData(ToIndex, userindex, 0, "||El soporte ha sido enviado con éxito. Gracias por utilizar nuestro sistema. Aguarde su respuesta." & FONTTYPE_furius)
        Exit Sub
        End If
'FuriusAO SISTEMA DE SOPORTE BASICO.


If UCase$(rdata) = "/RETAR" Then
If UserList(userindex).flags.EnTorneo Then Exit Sub

If RetoDesactivado Then Exit Sub
    If RetoEnCurso = True Then
        Call SendData(ToIndex, userindex, 0, "||Hay otro reto en curso" & FONTTYPE_CELESTE)
        Exit Sub
    End If
    If UserList(userindex).flags.TargetUser <> 0 Then

        If UserList(userindex).POS.Map <> 160 Then
            Call SendData(ToIndex, userindex, 0, "||Para retar a alguien debes estar en el mapa 160" & FONTTYPE_VENENO)
            Exit Sub
        End If
    Dim RetoTargetUser As Integer
    If RetoTargetUser <> 0 Then
    UserList(RetoTargetUser).flags.RetadoPor = 0
    UserList(RetoTargetUser).flags.Retado = 0
    End If
    RetoTargetUser = UserList(userindex).flags.TargetUser
    If RetoTargetUser = userindex Then Exit Sub
    UserList(RetoTargetUser).flags.Retado = 1
    UserList(RetoTargetUser).flags.RetadoPor = userindex
    UserList(userindex).flags.RetadoPor = RetoTargetUser
    Call SendData(ToIndex, userindex, 0, "||Has retado a " & UserList(RetoTargetUser).Name & FONTTYPE_CELESTE)
    Call SendData(ToIndex, RetoTargetUser, 0, "||" & UserList(userindex).Name & " te ha retado /ACEPTAR" & FONTTYPE_CELESTE)
    Else
    Call SendData(ToIndex, userindex, 0, "||Debes seleccionar a alguien para retar" & FONTTYPE_VENENO)
    End If
    Exit Sub
End If


If UCase$(rdata) = "/PAREJA" Then
If UserList(userindex).flags.EnTorneo Then Exit Sub
If UserList(userindex).flags.EnReto Then Exit Sub
If ParejasDesactivado Then Exit Sub
    If Reto2vs2EnCursO = True Then
        Call SendData(ToIndex, userindex, 0, "||Hay otro reto en curso" & FONTTYPE_CELESTE)
        Exit Sub
    End If
    
    If UserList(userindex).flags.TargetUser <> 0 Then

    Dim Reto2vs2TU As Integer
    
    Reto2vs2TU = UserList(userindex).flags.TargetUser
    If Reto2vs2TU = userindex Then Exit Sub
    
        If UserList(userindex).POS.Map <> 160 Then
            Call SendData(ToIndex, userindex, 0, "||Para pedir pareja debes estar en el mapa 160" & FONTTYPE_VENENO)
            Exit Sub
        End If
        
        If UserList(Reto2vs2TU).POS.Map <> 160 Then
            Call SendData(ToIndex, userindex, 0, "||Para pedir pareja debe estar en el mapa 160" & FONTTYPE_VENENO)
            Exit Sub
        End If
        
        If UserList(userindex).Stats.GLD < 100000 Then
            Call SendData(ToIndex, userindex, 0, "||Para pedir pareja debes tener por lo menos 100.000 monedas de oro" & FONTTYPE_VENENO)
            Exit Sub
        End If
    
    UserList(Reto2vs2TU).flags.Parejado = userindex
    'UserList(userindex).flags.Pareja = Reto2vs2TU
    Call SendData(ToIndex, userindex, 0, "||Le has pedido ser su pareja a " & UserList(Reto2vs2TU).Name & FONTTYPE_CELESTE)
    Call SendData(ToIndex, Reto2vs2TU, 0, "||" & UserList(userindex).Name & " te ha pedido ser su pareja /SIPAREJA" & FONTTYPE_CELESTE)
    Else
    Call SendData(ToIndex, userindex, 0, "||Debes seleccionar a alguien de pareja" & FONTTYPE_VENENO)
    End If
    Exit Sub
End If

If UCase$(rdata) = "/SIPAREJA" Then
If UserList(userindex).flags.EnTorneo Then Exit Sub
If UserList(userindex).flags.EnReto Then Exit Sub
If RetoDesactivado Then Exit Sub

If UserList(userindex).flags.Muerto = 1 Then Exit Sub
If UserList(userindex).POS.Map <> 160 Then Exit Sub
If UserList(userindex).flags.Parejado = 0 Then Exit Sub

If UserList(UserList(userindex).flags.Parejado).POS.Map <> 160 Then Exit Sub
If UserList(UserList(userindex).flags.Parejado).flags.Muerto = 1 Then Exit Sub
If UserList(UserList(userindex).flags.Parejado).flags.EnReto Then Exit Sub
If UserList(UserList(userindex).flags.Parejado).flags.Pareja > 0 Then Exit Sub



If UserList(userindex).Stats.GLD < 100000 Then
     Call SendData(ToIndex, userindex, 0, "||Para aceptar pareja debes tener por lo menos 100.000 monedas de oro" & FONTTYPE_VENENO)
     Exit Sub
End If

If UserList(UserList(userindex).flags.Parejado).Stats.GLD < 100000 Then
     Call SendData(ToIndex, userindex, 0, "||Para aceptar tu pareja debe tener por lo menos 100.000 monedas de oro" & FONTTYPE_VENENO)
     Exit Sub
End If

If Reto2vs2EnCursO Then
    Call SendData(ToIndex, userindex, 0, "||Hay otro reto 2vs2 en curso" & FONTTYPE_CELESTE)
    Exit Sub
End If

If UserList(UserList(userindex).flags.Parejado).flags.UserLogged = False Then Exit Sub


UserList(userindex).flags.EnReto = 1
UserList(UserList(userindex).flags.Parejado).flags.EnReto = 1
UserList(userindex).flags.Pareja = UserList(userindex).flags.Parejado
UserList(UserList(userindex).flags.Parejado).flags.Pareja = userindex

Call SendData(ToIndex, userindex, 0, "||Has aceptado ser la pareja de " & UserList(UserList(userindex).flags.Parejado).Name & FONTTYPE_BLANCO)
Call SendData(ToIndex, UserList(userindex).flags.Parejado, 0, "||" & UserList(userindex).Name & " ha aceptado ser tu pareja" & FONTTYPE_BLANCO)

UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 100000
UserList(UserList(userindex).flags.Parejado).Stats.GLD = UserList(UserList(userindex).flags.Parejado).Stats.GLD - 100000

If Pareja1.User1 > 0 Then
Pareja2.User1 = userindex
Pareja2.User2 = UserList(userindex).flags.Parejado
'COMENZAR DUELO
Call LlevarParejas
Reto2vs2EnCursO = True

Else
Pareja1.User1 = userindex
Pareja1.User2 = UserList(userindex).flags.Parejado
    If Pareja2.User1 > 0 Then
    
    'COMENZAR DUELO
    Call LlevarParejas
    Reto2vs2EnCursO = True
    
    End If
End If



Exit Sub

End If




If UCase$(rdata) = "/ACEPTAR" Then
If ParejasDesactivado Then Exit Sub
If UserList(userindex).flags.EnReto Then Exit Sub
If UserList(userindex).flags.EnTorneo Then Exit Sub
If UserList(userindex).flags.Retado = 0 Then Exit Sub
If UserList(userindex).flags.Muerto = 1 Then Exit Sub
If UserList(userindex).POS.Map <> 160 Then Exit Sub
If UserList(userindex).flags.Pareja > 0 Then Exit Sub

If UserList(UserList(userindex).flags.RetadoPor).POS.Map <> 160 Then Exit Sub
If UserList(UserList(userindex).flags.RetadoPor).flags.Muerto = 1 Then Exit Sub
If UserList(UserList(userindex).flags.RetadoPor).flags.Pareja > 0 Then Exit Sub
If UserList(UserList(userindex).flags.RetadoPor).flags.EnReto > 0 Then Exit Sub

If RetoEnCurso Then
    Call SendData(ToIndex, userindex, 0, "||Hay otro reto en curso" & FONTTYPE_CELESTE)
    Exit Sub
End If

If UserList(UserList(userindex).flags.RetadoPor).flags.UserLogged = False Then Exit Sub
CuentaRegresiva = 4
GMCuenta = 170
Call WarpUserChar(userindex, 170, 38, 48, True)
Call WarpUserChar(UserList(userindex).flags.RetadoPor, 170, 51, 55, True)
UserList(userindex).flags.EnReto = 1
UserList(UserList(userindex).flags.RetadoPor).flags.EnReto = 1
TiempoReto = 150
RetoJ(1) = userindex
RetoJ(2) = UserList(userindex).flags.RetadoPor
RetoEnCurso = True
Exit Sub
End If



If UCase$(Left$(rdata, 10)) = "/VERPENAS " Then
Dim NPen As Integer
Dim UserNom As String
Dim PenasFinal As String
UserNom = Trim(Right$(rdata, Len(rdata) - 10))
If UserNom <> UserList(userindex).Name And UserList(userindex).flags.Privilegios < 1 Then Exit Sub
NPen = val(GetVar(CharPath & "/" & UCase$(UserNom) & ".chr", "PENAS", "CANTIDAD"))
PenasFinal = "PENAS: " & NPen
    If NPen > 0 Then
    Dim Bcl As Integer
    For Bcl = 1 To NPen
    PenasFinal = PenasFinal & "PENA" & Bcl & ":" & GetVar(CharPath & "/" & UCase$(UserNom) & ".chr", "PENAS", "PENA" & Bcl) & vbCrLf
    DoEvents
    Next Bcl
    Call SendData(ToIndex, userindex, 0, "||" & PenasFinal & FONTTYPE_BLANCO)
    Else
    Call SendData(ToIndex, userindex, 0, "||NO POSEE PENAS" & FONTTYPE_BLANCO)
    End If
End If




If UCase$(rdata) = "/HOGAR" Then
    If Not ModoQuest Then Exit Sub
    If UserList(userindex).flags.Muerto = 0 Then Exit Sub
    If UserList(userindex).POS.Map = ULLATHORPE.Map Then Exit Sub
    Call WarpUserChar(userindex, ULLATHORPE.Map, ULLATHORPE.x, ULLATHORPE.Y, True)
    Exit Sub
End If

If UCase$(Left$(rdata, 12)) = "/MERCENARIO " Then
    rdata = Right$(rdata, Len(rdata) - 12)
    If Not ModoQuest Then Exit Sub
    If UserList(userindex).flags.Privilegios > 0 Then Exit Sub
    Select Case UCase$(rdata)
        Case "ALIANZA"
            tInt = 1
        Case "LORD THEK"
            tInt = 2
        Case Else
            Call SendData(ToIndex, userindex, 0, "||La estructura del comando es /MERCENARIO ALIANZA o /MERCENARIO LORD THEK." & FONTTYPE_furius)
            Exit Sub
    End Select
    
    Select Case UserList(userindex).Faccion.BandoOriginal
        Case Neutral
            If UserList(userindex).Faccion.Bando <> Neutral Then
                Call SendData(ToIndex, userindex, 0, "||Ya eres mercenario para " & ListaBandos(UserList(userindex).Faccion.Bando) & "." & FONTTYPE_furius)
                Exit Sub
            End If
        
        Case Else
            Select Case UserList(userindex).Faccion.Bando
                Case Neutral
                    If tInt = UserList(userindex).Faccion.BandoOriginal Then
                        Call SendData(ToIndex, userindex, 0, "||" & ListaBandos(tInt) & " no acepta desertores entre sus filas." & FONTTYPE_furius)
                        Exit Sub
                    End If
            
                Case UserList(userindex).Faccion.BandoOriginal
                    Call SendData(ToIndex, userindex, 0, "||Ya perteneces a " & ListaBandos(UserList(userindex).Faccion.Bando) & ", no puedes ofrecerte como mercenario." & FONTTYPE_furius)
                    Exit Sub
        
                Case Else
                    Call SendData(ToIndex, userindex, 0, "||Ya eres mercenario para " & ListaBandos(UserList(userindex).Faccion.Bando) & "." & FONTTYPE_furius)
                    Exit Sub
            End Select
    End Select
    Call SendData(ToIndex, userindex, 0, "||¡" & ListaBandos(tInt) & " te ha aceptado como un mercenario entre sus filas!" & FONTTYPE_furius)
    UserList(userindex).Faccion.Bando = tInt
    Call UpdateUserChar(userindex)
    Exit Sub
End If

If UserList(userindex).flags.Quest Then
    If UCase$(Left$(rdata, 3)) = "/M " Then
        rdata = Right$(rdata, Len(rdata) - 3)
        If Len(rdata) = 0 Then Exit Sub
        Select Case UserList(userindex).Faccion.Bando
            Case Real
                tStr = FONTTYPE_ARMADA
            Case Caos
                tStr = FONTTYPE_CAOS
        End Select
        Call SendData(ToAll, 0, 0, "||" & rdata & tStr)
        Exit Sub
    ElseIf UCase$(rdata) = "/TELEPLOC" Then
        Call WarpUserChar(userindex, UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY, True)
        Exit Sub
    ElseIf UCase$(rdata) = "/TRAMPA" Then
        Call ActivarTrampa(userindex)
        Exit Sub
    End If
End If

If UserList(userindex).flags.PuedeDenunciar Or UserList(userindex).flags.Privilegios > 0 Then
    If UCase$(Left$(rdata, 11)) = "/DENUNCIAS " Then
        rdata = Right$(rdata, Len(rdata) - 11)
        tIndex = NameIndex(rdata)
        
        If tIndex > 0 Then
            Call SendData(ToIndex, userindex, 0, "||Denuncias por cheat: " & UserList(tIndex).flags.Denuncias & "." & FONTTYPE_INFO)
            Call SendData(ToIndex, userindex, 0, "||Denuncias por insultos: " & UserList(tIndex).flags.DenunciasInsultos & "." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, userindex, 0, "1A")
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 6)) = "/DENC " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        tIndex = NameIndex(rdata)
        
        If tIndex > 0 Then
            UserList(tIndex).flags.Denuncias = UserList(tIndex).flags.Denuncias + 1
            Call SendData(ToIndex, userindex, 0, "||Sumaste una denuncia por cheat a " & UserList(tIndex).Name & ". El usuario tiene acumuladas " & UserList(tIndex).flags.Denuncias & " denuncias." & FONTTYPE_INFO)
            Call LogGM(UserList(userindex).Name, "Sumo una denuncia por cheat a " & UserList(tIndex).Name & ".", UserList(userindex).flags.Privilegios = 1)
        Else
            If Not ExistePersonaje(rdata) Then
                Call SendData(ToIndex, userindex, 0, "||El personaje está offline y no se encuentra en la base de datos." & FONTTYPE_INFO)
                Exit Sub
            End If
            Call LogGM(UserList(userindex).Name, "Sumo una denuncia por cheat a " & rdata & ".", UserList(userindex).flags.Privilegios = 1)
            Call SendData(ToIndex, userindex, 0, "||Sumaste una denuncia por cheat a " & rdata & ". El usuario tiene acumuladas " & SumarDenuncia(rdata, 1) & " denuncias." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 6)) = "/DENI " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        tIndex = NameIndex(rdata)
        
        If tIndex > 0 Then
            UserList(tIndex).flags.DenunciasInsultos = UserList(tIndex).flags.DenunciasInsultos + 1
            Call SendData(ToIndex, userindex, 0, "||Sumaste una denuncia por insultos a " & UserList(tIndex).Name & ". El usuario tiene acumuladas " & UserList(tIndex).flags.DenunciasInsultos & " denuncias." & FONTTYPE_INFO)
            Call LogGM(UserList(userindex).Name, "Sumo una denuncia por insultos a " & UserList(tIndex).Name & ".", UserList(userindex).flags.Privilegios = 1)
        Else
            If Not ExistePersonaje(rdata) Then
                Call SendData(ToIndex, userindex, 0, "||El personaje está offline y no se encuentra en la base de datos." & FONTTYPE_INFO)
                Exit Sub
            End If
            Call LogGM(UserList(userindex).Name, "Sumo una denuncia por insultos a " & rdata & ".", UserList(userindex).flags.Privilegios = 1)
            Call SendData(ToIndex, userindex, 0, "||Sumaste una denuncia por insultos a " & rdata & ". El usuario tiene acumuladas " & SumarDenuncia(rdata, 2) & " denuncias." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
End If

If UserList(userindex).flags.Privilegios = 0 Then Exit Sub

If UCase$(rdata) = "/TIEMPOMOMIA" Then
Call SendData(ToIndex, userindex, 0, "||Tiempo momia: " & TiempoMomia & FONTTYPE_VENENO)
Exit Sub
End If




If UCase$(rdata) = "/DAMESOS" Then
Dim LstU As String
    
    If Soportes.Count = 0 Then
        Call SendData(ToIndex, userindex, 0, "||No hay soportes para ver." & FONTTYPE_BLANCO)
        Exit Sub
    End If

    For i = 1 To Soportes.Count
        LstU = LstU & "@" & Soportes.Item(i)
        Debug.Print Soportes.Item(i)
        DoEvents
    Next i

    LstU = Soportes.Count & LstU

    LstU = "SHWSOP@" & LstU
    Call SendData(ToIndex, userindex, 0, LstU)
    
End If

If UCase$(Left$(rdata, 7)) = "/BORSO " Then
rdata = Right$(rdata, Len(rdata) - 7)
Call WriteVar(CharPath & UCase$(rdata) & ".chr", "STATS ", "Soporte", "")
Call WriteVar(CharPath & UCase$(rdata) & ".chr", "STATS ", "Respuesta", "")
For i = 1 To Soportes.Count
If UCase$(Soportes.Item(i)) = UCase$(rdata) Then
    Soportes.Remove (i)
    Exit For
End If
DoEvents
Next i
Call SendData(ToIndex, userindex, 0, "||Soporte y respuesta borrados con éxito" & FONTTYPE_BLANCO)
Exit Sub
End If


If UCase$(Left$(rdata, 7)) = "/SOSDE " Then
rdata = Right$(rdata, Len(rdata) - 7)

Dim SosDe As String
SosDe = GetVar(CharPath & UCase$(rdata) & ".chr", "STATS", "Soporte")


    If Len(SosDe) > 0 Then
        Call SendData(ToIndex, userindex, 0, "SOPODE" & SosDe)
    Else
        Call SendData(ToIndex, userindex, 0, "||Error. Soporte no encontrado" & FONTTYPE_BLANCO)
    End If


End If

If UCase$(Left$(rdata, 7)) = "/RESOS " Then
rdata = Right$(rdata, Len(rdata) - 7)
Dim Persona, Respuesta As String
Persona = ReadField$(1, rdata, Asc(";")) 'GetVar(CharPath & UCase$(rdata) & ".chr", "STATS", "Soporte")
Respuesta = Replace(ReadField$(2, rdata, Asc(";")), Chr$(13) & Chr$(10), Chr(32))
If Len(Persona) = 0 Or Len(Respuesta) = 0 Then
    Call SendData(ToIndex, userindex, 0, "||Error en la respuesta" & FONTTYPE_BLANCO)
    Exit Sub
End If

Call WriteVar(CharPath & UCase$(Persona) & ".chr", "STATS", "Respuesta", Respuesta)
Call WriteVar(CharPath & UCase$(Persona) & ".chr", "STATS", "Soporte", GetVar(CharPath & UCase$(Persona) & ".chr", "STATS", "Soporte") & "0k1")


tIndex = NameIndex(Persona)
If tIndex > 0 Then
    Call SendData(ToIndex, tIndex, 0, "||Tu soporte ha sido respondido." & FONTTYPE_furius)
    Call SendData(ToIndex, tIndex, 0, "TENSO")
End If
    
Call SendData(ToIndex, userindex, 0, "||Soporte respondido con éxito" & FONTTYPE_BLANCO)
    For i = 1 To Soportes.Count
    Debug.Print Soportes.Item(1)
    
        If UCase$(Soportes.Item(i)) = UCase$(Persona) Then
            Soportes.Remove (i)
            Exit For
        End If
        DoEvents
    Next i


End If






If UCase$(Left$(rdata, 4)) = "/GO " Then
    rdata = Right$(rdata, Len(rdata) - 4)
    mapa = val(ReadField(1, rdata, 32))
    If Not MapaValido(mapa) Then Exit Sub
    'If UserList(userindex).flags.Privilegios = 1 And MapInfo(mapa).Pk Then Exit Sub
    Call WarpUserChar(userindex, mapa, 50, 50, True)
    Call SendData(ToIndex, userindex, 0, "2B" & UserList(userindex).Name)
    Call LogGM(UserList(userindex).Name, "Transporto a " & UserList(userindex).Name & " hacia " & "Mapa" & mapa & " X:" & x & " Y:" & Y, (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/PTORNEO" Then
Call SendData(ToIndex, userindex, 0, "PTOR")
Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/NPC " Then
            If UserList(userindex).flags.TargetNpc > 0 Then
                tStr = Right$(rdata, Len(rdata) - 5)
                Call SendData(ToNPCArea, UserList(userindex).flags.TargetNpc, Npclist(UserList(userindex).flags.TargetNpc).POS.Map, "||" & vbGreen & "°" & tStr & "°" & CStr(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
                Call LogGM(UserList(userindex).Name, "Dijo x NPC: " & tStr, False)
            Else
                Call SendData(ToIndex, userindex, 0, "||Debes seleccionar el NPC por el que quieres hablar antes de usar este comando" & FONTTYPE_INFO)
            End If
        Exit Sub
End If



If UCase$(Left$(rdata, 5)) = "/SUM " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(userindex).flags.Privilegios < UserList(tIndex).flags.Privilegios And UserList(tIndex).flags.AdminInvisible = 1 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(userindex).flags.Privilegios = 1 And UserList(tIndex).POS.Map <> UserList(userindex).POS.Map Then Exit Sub
    
    Call SendData(ToIndex, userindex, 0, "%Z" & UserList(tIndex).Name)
    Call WarpUserChar(tIndex, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y + 1, True)
    
    Call LogGM(UserList(userindex).Name, "/SUM " & UserList(tIndex).Name & " Map:" & UserList(userindex).POS.Map & " X:" & UserList(userindex).POS.x & " Y:" & UserList(userindex).POS.Y, False)
    Exit Sub
End If

If UCase$(rdata) = "/INVISIBLE" Then
    Call DoAdminInvisible(userindex)
   ' Call LogGM(UserList(userindex).Name, "/INVISIBLE", (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(rdata) = "/TELEPLOC" Then
    Call WarpUserChar(userindex, UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY, True)
  '  Call LogGM(UserList(userindex).Name, "/TELEPLOC a x:" & UserList(userindex).flags.TargetX & " Y:" & UserList(userindex).flags.TargetY & " Map:" & UserList(userindex).POS.Map, (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/STAFF " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Call LogGM(UserList(userindex).Name, "Mensaje a Gms:" & rdata, (UserList(userindex).flags.Privilegios = 1))
    If Len(rdata) > 0 Then
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & "> " & rdata & "~0~191~255~0~0")
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/ECHAR " Then

    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(rdata)

    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1E")
        Exit Sub
    End If
    
    If tIndex = userindex Then Exit Sub
    
    If UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
        Call SendData(ToIndex, userindex, 0, "1F")
        Exit Sub
    End If
        
    Call SendData(ToAdmins, 0, 0, "%U" & UserList(userindex).Name & "," & UserList(tIndex).Name)
    Call LogGM(UserList(userindex).Name, "Echo a " & UserList(tIndex).Name, False)
    Call CloseSocket(tIndex)
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/IRA " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    If ((UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios And UserList(tIndex).flags.AdminInvisible = 1)) Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If

    If UserList(tIndex).flags.AdminInvisible And Not UserList(userindex).flags.AdminInvisible Then Call DoAdminInvisible(userindex)

    Call WarpUserChar(userindex, UserList(tIndex).POS.Map, UserList(tIndex).POS.x + 1, UserList(tIndex).POS.Y + 1, True)
    
    Call LogGM(UserList(userindex).Name, "/IRA " & UserList(tIndex).Name & " Mapa:" & UserList(tIndex).POS.Map & " X:" & UserList(tIndex).POS.x & " Y:" & UserList(tIndex).POS.Y, (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(rdata) = "/TRABAJANDO" Then
    For LoopC = 1 To LastUser
        If Len(UserList(LoopC).Name) > 0 And UserList(LoopC).flags.Trabajando Then
            DummyInt = DummyInt + 1
            tStr = tStr & UserList(LoopC).Name & ", "
        End If
    Next
    If Len(tStr) > 0 Then
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, userindex, 0, "||Usuarios trabajando: " & tStr & FONTTYPE_INFO)
        Call SendData(ToIndex, userindex, 0, "||Número de usuarios trabajando: " & DummyInt & "." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, userindex, 0, "%)")
    End If
    Exit Sub
End If



If UCase$(Left$(rdata, 11)) = "/SILENCIAR " Then
   rdata = Right$(rdata, Len(rdata) - 11)
   
    Arg1 = ReadField(1, rdata, 64)
    Name = ReadField(3, rdata, 64)
    i = val(ReadField(2, rdata, 64))
    
      If Len(Arg1) = 0 Or Len(Name) = 0 Or i = 0 Then
        Call SendData(ToIndex, userindex, 0, "||La estructura del comando es /SILENCIAR CAUSA@MINS@NICK." & FONTTYPE_furius)
        Exit Sub
    End If
    
   
    tIndex = NameIndex(Name)
    If tIndex Then
    Call LogPENA(UserList(tIndex).Name, "SILENCIADO. Motivo: " & Arg1 & ".. Tiempo: " & i, userindex)
    If i > 60 Then i = 60
    UserList(tIndex).flags.Silenciado = i
     Call SendData(ToIndex, tIndex, 0, "||" & "Has sido silenciado por " & Arg1 & " durante los próximos " & i & " minutos." & " GM: " & UserList(userindex).Name & FONTTYPE_BLANCO)
     Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Silencio a " & UserList(tIndex).Name & " por " & Arg1 & " durante los próximos " & i & " minutos." & FONTTYPE_furius)
    Call LogGM(UserList(userindex).Name, "/SILENCIAR a " & UserList(tIndex).Name & " por " & Arg1 & " durante los próximos " & i & " minutos.", (UserList(userindex).flags.Privilegios = 1))
  ' Call SendData(ToAll, 0, 0, "||Servidor>" & UserList(tIndex).Name & " ha sido silenciado por " & i & " minutos." & FONTTYPE_BLANCO)
    Else
    Call SendData(ToIndex, userindex, 0, "||El usuario está offline" & FONTTYPE_BLANCO)
    End If
Exit Sub
End If






If UCase$(Left$(rdata, 9)) = "/LOGPENA " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Call LogPENA(UserList(UserList(userindex).flags.TargetUser).Name, "LogPena: " & rdata, userindex)
    Call SendData(ToIndex, userindex, 0, "||Pena grabada con éxito" & FONTTYPE_INFO)
    Exit Sub
End If


'PenaMinar

If UCase$(Left$(rdata, 8)) = "/CARCEL " Then
    
rdata = Right$(rdata, Len(rdata) - 8)
Arg1 = ReadField(1, rdata, 64)
Name = ReadField(2, rdata, 64)
i = val(ReadField(3, rdata, 64))
 
If Len(Arg1) = 0 Or Len(Name) = 0 Or i = 0 Then
Call SendData(ToIndex, userindex, 0, "||La estructura del comando es /CARCEL CAUSA@NICK@CANTIDAD." & FONTTYPE_furius)
Exit Sub
End If
tIndex = NameIndex(Name)
If tIndex <= 0 Then
    Call SendData(ToIndex, userindex, 0, "1A")
    Exit Sub
End If
    
If UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
    Call SendData(ToIndex, userindex, 0, "1B")
    Exit Sub
End If
    
If i > 999999 Then
    Call SendData(ToIndex, userindex, 0, "1C")
    Exit Sub
End If
UserList(tIndex).Counters.PenaMinar = i
UserList(tIndex).flags.Encarcelado = 1
UserList(tIndex).Counters.Pena = Timer

Dim ItemVaria As Obj
ItemVaria.Amount = 1
ItemVaria.OBJIndex = 187
If Not MeterItemEnInventario(tIndex, ItemVaria) Then
Call SendData(ToIndex, userindex, 0, "||NO SE LE PUDO DAR EL PICO DE MINERO" & FONTTYPE_BLANCO)
End If
  
ItemVaria.OBJIndex = 31
If Not MeterItemEnInventario(tIndex, ItemVaria) Then
Call SendData(ToIndex, userindex, 0, "||NO SE LE PUDO DAR LA ROPA" & FONTTYPE_BLANCO)
End If
  
ItemVaria.OBJIndex = 240
If Not MeterItemEnInventario(tIndex, ItemVaria) Then
Call SendData(ToIndex, userindex, 0, "||NO SE LE PUDO DAR LA ROPA" & FONTTYPE_BLANCO)
End If
  
Call WarpUserChar(tIndex, Prision.Map, Prision.x, Prision.Y, True)
Call LogPENA(UserList(tIndex).Name, "ENCARCELADO. Motivo: " & Arg1 & ".. Cantidad: " & i, userindex)
Call SendData(ToIndex, tIndex, 0, "||" & "Has sido encarcelado por " & Arg1 & ". Para salir debes minar " & i & " piezas de hierro." & " GM: " & UserList(userindex).Name & FONTTYPE_BLANCO)
Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Encarcelo a " & UserList(tIndex).Name & " por " & Arg1 & " por " & i & " piedras." & FONTTYPE_furius)
Call LogGM(UserList(userindex).Name, "/CARCEL a " & UserList(tIndex).Name & " por " & Arg1 & " por " & i & " piedras.", (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If










If UCase$(Left$(rdata, 14)) = "/CARCELTIEMPO " Then
    
    rdata = Right$(rdata, Len(rdata) - 14)
    
   Arg1 = ReadField(1, rdata, 64)
    Name = ReadField(2, rdata, 64)
    i = val(ReadField(3, rdata, 64))
    
   If Len(Arg1) = 0 Or Len(Name) = 0 Or i = 0 Then
     Call SendData(ToIndex, userindex, 0, "||La estructura del comando es /CARCELTIEMPO CAUSA@NICK@MINUTOS." & FONTTYPE_furius)
     Exit Sub
    End If
    
    tIndex = NameIndex(Name)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
        Call SendData(ToIndex, userindex, 0, "1B")
        Exit Sub
    End If
    
    If i > 120 Then
       Call SendData(ToIndex, userindex, 0, "1C")
        Exit Sub
    End If
UserList(tIndex).Counters.TiempoPena = 60 * i
UserList(tIndex).flags.Encarcelado = 1
UserList(tIndex).Counters.Pena = Timer
Call WarpUserChar(tIndex, Prision.Map, Prision.x, Prision.Y, True)
Call LogPENA(UserList(tIndex).Name, "ENCARCELADO. Motivo: " & Arg1 & ".. Tiempo: " & i, userindex)
Call SendData(ToIndex, tIndex, 0, "||" & "Has sido encarcelado por " & Arg1 & " durante los próximos " & i & " minutos." & " GM: " & UserList(userindex).Name & FONTTYPE_BLANCO)
Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Encarcelo a " & UserList(tIndex).Name & " por " & Arg1 & " durante los próximos " & i & " minutos." & FONTTYPE_furius)
Call LogGM(UserList(userindex).Name, "/CARCEL a " & UserList(tIndex).Name & " por " & Arg1 & " durante los próximos " & i & " minutos.", (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UserList(userindex).flags.Privilegios < 2 Then Exit Sub

If UCase$(Left$(rdata, 4)) = "/REM" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    Call LogGM(UserList(userindex).Name, "Comentario: " & rdata, (UserList(userindex).flags.Privilegios = 1))
    Call SendData(ToIndex, userindex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
    Exit Sub
End If



If UCase$(Left$(rdata, 5)) = "/HORA" Then
    Call LogGM(UserList(userindex).Name, "Hora.", (UserList(userindex).flags.Privilegios = 1))
    rdata = Right$(rdata, Len(rdata) - 5)
    Call SendData(ToAll, 0, 0, "||Hora: " & Time & " " & Date & "~0~191~255~0~0")
    Exit Sub
End If

If UCase$(rdata) = "/ONLINEGM" Then
        For LoopC = 1 To LastUser
            If Len(UserList(LoopC).Name) > 0 Then
                If UserList(LoopC).flags.Privilegios > 0 And (UserList(LoopC).flags.Privilegios <= UserList(userindex).flags.Privilegios Or UserList(LoopC).flags.AdminInvisible = 0) Then
                    tStr = tStr & UserList(LoopC).Name & ", "
                End If
            End If
        Next
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, userindex, 0, "||" & tStr & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, userindex, 0, "%P")
        End If
        Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/DONDE " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios And UserList(tIndex).flags.AdminInvisible = 1 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    Call SendData(ToIndex, userindex, 0, "||Ubicacion de " & UserList(tIndex).Name & ": " & UserList(tIndex).POS.Map & ", " & UserList(tIndex).POS.x & ", " & UserList(tIndex).POS.Y & "." & FONTTYPE_INFO)
    Call LogGM(UserList(userindex).Name, "/Donde", (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/NENE " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    If MapaValido(val(rdata)) Then
        Call SendData(ToIndex, userindex, 0, "NENE" & NPCHostiles(val(rdata)))
        Call LogGM(UserList(userindex).Name, "Numero enemigos en mapa " & rdata, (UserList(userindex).flags.Privilegios = 1))
    End If
    Exit Sub
End If

If UCase$(rdata) = "/VENTAS" Then
    Call SendData(ToIndex, userindex, 0, "/X" & DineroTotalVentas & "," & NumeroVentas)
    Exit Sub
End If



If UCase$(rdata) = "/DESCONGELAR" Then
    Call Congela(True)
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/VIGILAR " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    tIndex = NameIndex(rdata)
    If tIndex > 0 Then
        If tIndex = userindex Then
            Call SendData(ToIndex, userindex, 0, "||No puedes vigilarte a ti mismo." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(tIndex).flags.Privilegios >= UserList(userindex).flags.Privilegios Then
            Call SendData(ToIndex, userindex, 0, "||No puedes vigilar a alguien con igual o mayor jerarquia que tú." & FONTTYPE_INFO)
            Exit Sub
        End If
        If YaVigila(tIndex, userindex) Then
            Call SendData(ToIndex, userindex, 0, "||Dejaste de vigilar a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
            If Not EsVigilado(tIndex) Then Call SendData(ToIndex, tIndex, 0, "VIG")
            Exit Sub
        End If
        If Not EsVigilado(tIndex) Then Call SendData(ToIndex, tIndex, 0, "VIG")
        Call SendData(ToIndex, userindex, 0, "||Estás vigilando a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
        For i = 1 To 10
            If UserList(tIndex).flags.Espiado(i) = 0 Then
                UserList(tIndex).flags.Espiado(i) = userindex
                Exit For
            End If
        Next
        If i = 11 Then
            Call SendData(ToIndex, userindex, 0, "||Demasiados GM's están vigilando a este usuario." & FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        Call SendData(ToIndex, userindex, 0, "1A")
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/TELEP " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    mapa = val(ReadField(2, rdata, 32))
    If Not MapaValido(mapa) Then Exit Sub
    Name = ReadField(1, rdata, 32)
    If Len(Name) = 0 Then Exit Sub
    If UCase$(Name) <> "YO" Then
        If UserList(userindex).flags.Privilegios = 1 Then
            Exit Sub
        End If
        tIndex = NameIndex(Name)
    Else
        tIndex = userindex
    End If
    x = val(ReadField(3, rdata, 32))
    Y = val(ReadField(4, rdata, 32))
    If Not InMapBounds(x, Y) Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    If UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios And UserList(userindex).flags.AdminInvisible = 1 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    Call WarpUserChar(tIndex, mapa, x, Y, True)
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(userindex).Name & " te ha transportado." & FONTTYPE_INFO)
    Call LogGM(UserList(userindex).Name, "Transporto a " & UserList(tIndex).Name & " hacia " & "Mapa" & mapa & " X:" & x & " Y:" & Y, (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If
'If UCase$(Left$(rdata, 4)) = "/GO " Then
'    rdata = Right$(rdata, Len(rdata) - 4)
'    mapa = val(ReadField(1, rdata, 32))
'    If Not MapaValido(mapa) Then Exit Sub
'    Call WarpUserChar(userindex, mapa, 50, 50, True)
'    Call SendData(ToIndex, userindex, 0, "2B" & UserList(userindex).Name)
'    Call LogGM(UserList(userindex).Name, "Transporto a " & UserList(userindex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(userindex).flags.Privilegios = 1))
'    Exit Sub
'End If




If UCase$(Left$(rdata, 6)) = "/RACC " Then
   rdata = val(Right$(rdata, Len(rdata) - 6))
      NumNPC = val(GetVar(App.Path & "\Dat\NPCs-HOSTILES.dat", "INIT", "NumNPCs")) + 500
   If rdata < 0 Or rdata > NumNPC Then
    Call SendData(ToIndex, userindex, 0, "||La criatura no existe." & FONTTYPE_INFO)
Else
   Call SpawnNpc(val(rdata), UserList(userindex).POS, True, True)
    Call LogGM(UserList(userindex).Name, rdata, False)
   End If
   Exit Sub
End If


If UCase$(rdata) = "/OMAP" Then
    For LoopC = 1 To MapInfo(UserList(userindex).POS.Map).NumUsers
        If UserList(MapInfo(UserList(userindex).POS.Map).userindex(LoopC)).flags.Privilegios <= UserList(userindex).flags.Privilegios Then
            tStr = tStr & UserList(MapInfo(UserList(userindex).POS.Map).userindex(LoopC)).Name & ","
        End If
    Next
    If Len(tStr) > 0 Then
        tStr = Left$(tStr, Len(tStr) - 1)
        Call SendData(ToIndex, userindex, 0, "||Usuarios en este mapa: " & tStr & "." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, userindex, 0, "%R")
    End If
    Exit Sub
End If

If UCase$(rdata) = "/PANELGM" Then
    Call SendData(ToIndex, userindex, 0, "PGM3")
    Exit Sub
End If

If UCase$(rdata) = "/CMAP" Then
    If MapInfo(UserList(userindex).POS.Map).NumUsers Then
        Call SendData(ToIndex, userindex, 0, "||Hay " & MapInfo(UserList(userindex).POS.Map).NumUsers & " usuarios en este mapa." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, userindex, 0, "%R")
    End If

    Exit Sub
End If


If UCase$(rdata) = "/TORNEO" Then
    If EnTorneo = 0 Then
        EnTorneo = 1
        If FileExist(App.Path & "/logs/torneo.log", vbNormal) Then Kill (App.Path & "/logs/torneo.log")
        Call SendData(ToIndex, userindex, 0, "||Has activado el torneo" & FONTTYPE_INFO)
    Else
        EnTorneo = 0
        Call SendData(ToIndex, userindex, 0, "||Has desactivado el torneo" & FONTTYPE_INFO)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/ACC " Then
   rdata = val(Right$(rdata, Len(rdata) - 5))
   NumNPC = val(GetVar(App.Path & "\Dat\NPCs-HOSTILES.dat", "INIT", "NumNPCs")) + 500
   If rdata < 0 Or rdata > NumNPC Then
       Call SendData(ToIndex, userindex, 0, "||La criatura no existe." & FONTTYPE_INFO)

Else
   Call SpawnNpc(val(rdata), UserList(userindex).POS, True, False)
    Call LogGM(UserList(userindex).Name, rdata, False)

   End If
   Exit Sub
End If

If UCase$(Left$(rdata, 3)) = "/CT" Then
    
    rdata = Right$(rdata, Len(rdata) - 4)
    Call LogGM(UserList(userindex).Name, "/CT: " & rdata, False)
    mapa = ReadField(1, rdata, 32)
    x = ReadField(2, rdata, 32)
    Y = ReadField(3, rdata, 32)
    
    If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).OBJInfo.OBJIndex Then
        Exit Sub
    End If
    If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).TileExit.Map Then
        Exit Sub
    End If
    If Not MapaValido(mapa) Or Not InMapBounds(x, Y) Then Exit Sub
    
    Dim ET As Obj
    ET.Amount = 1
    ET.OBJIndex = Teleport
    
    Call MakeObj(ToMap, 0, UserList(userindex).POS.Map, ET, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1)
    MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).TileExit.Map = mapa
    MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).TileExit.x = x
    MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).TileExit.Y = Y
    
    Exit Sub
End If

'FIXIT: Reexmplazar la función 'Left' con la función 'Left$'.                               FixIT90210ae-R9757-R1B8ZE
If UCase(Left(rdata, 3)) = "/DT" Then
    '/dt
    Call LogGM(UserList(userindex).Name, "/DT", False)
    
    mapa = UserList(userindex).flags.TargetMap
    x = UserList(userindex).flags.TargetX
    Y = UserList(userindex).flags.TargetY
        Call EraseObj(ToMap, 0, mapa, MapData(mapa, x, Y).OBJInfo.Amount, mapa, x, Y)
        Call EraseObj(ToMap, 0, MapData(mapa, x, Y).TileExit.Map, 1, MapData(mapa, x, Y).TileExit.Map, MapData(mapa, x, Y).TileExit.x, MapData(mapa, x, Y).TileExit.Y)
        MapData(mapa, x, Y).TileExit.Map = 0
        MapData(mapa, x, Y).TileExit.x = 0
        MapData(mapa, x, Y).TileExit.Y = 0
     Exit Sub
     End If

 If UCase$(Left$(rdata, 11)) = "/FORCEMIDI " Then
    rdata = Right$(rdata, Len(rdata) - 11)
    If Not IsNumeric(rdata) Then
        Exit Sub
    Else
        Call SendData(ToAll, 0, 0, "||Broadcast Musica: " & rdata & FONTTYPE_INFO)
        Call SendData(ToAll, 0, 0, "TM" & rdata)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & "mando un broadcast musica " & FONTTYPE_INFO)
    End If
End If

If UCase$(Left$(rdata, 10)) = "/FORCEWAV " Then
    rdata = Right$(rdata, Len(rdata) - 10)
    If Not IsNumeric(rdata) Then
        Exit Sub
    Else
    'leito
     Call SendData(ToAll, 0, 0, "||Broadcast Musica: " & rdata & FONTTYPE_INFO)
     'leito he guacho recatate gil
        Call SendData(ToAll, 0, 0, "TW" & rdata)
        Call SendData(ToAdmins, 0, 0, "|| " & UserList(userindex).Name & "mando un broadcast wav: " & FONTTYPE_INFO)
    End If
End If


If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
    Call LogGM(UserList(userindex).Name, "/BLOQ", False)
    rdata = Right$(rdata, Len(rdata) - 5)
    If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).Blocked = 0 Then
        MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).Blocked = 1
        Call Bloquear(ToMap, userindex, UserList(userindex).POS.Map, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y, 1)
    Else
        MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).Blocked = 0
        Call Bloquear(ToMap, userindex, UserList(userindex).POS.Map, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y, 0)
    End If
    Exit Sub
End If



    




If UCase(rdata) = "/NOCHE" Then
        Call SendData(ToAll, 0, 0, "NUB" & IIf(Anochecer, 1, 0))
        Anochecer = Not Anochecer
        Exit Sub
End If

If UCase(rdata) = "/TARDE" Then
        Call SendData(ToAll, 0, 0, "TAR" & IIf(Atardecer, 1, 0))
       Atardecer = Not Atardecer
        Exit Sub
End If

If UCase(rdata) = "/MAÑANA" Then
        Call SendData(ToAll, 0, 0, "MAÑ" & IIf(Amanecer, 1, 0))
        Amanecer = Not Amanecer
        Exit Sub
End If



If UCase$(rdata) = "/LIMPIAROFERTAS" Then
    Ofertas.Reset
    Call SendData(ToIndex, userindex, 0, "||ColaBuscados Resetada" & FONTTYPE_BLANCO)
Exit Sub
End If



If UCase$(rdata) = "/LIMPIARMAPAS" Then
'Call LimpiarMapas
    Call LogGM(UserList(userindex).Name, "/LIMPIARMAPAS", (UserList(userindex).flags.Privilegios = 1))
frmMain.TLimpiarMapas = 62
Exit Sub
End If

If UCase$(rdata) = "/INIS" Then
Call LoadSini
Call SendData(ToIndex, userindex, 0, "||Inis reloaded" & FONTTYPE_VENENO)
    
Exit Sub
End If

' MapInfo(UserList(userindex).POS.Map).QuestMod
If UCase(rdata) = "/MODOQUESTMAP" Then
MapInfo(UserList(userindex).POS.Map).QuestMod = Not MapInfo(UserList(userindex).POS.Map).QuestMod
    Call SendData(ToIndex, userindex, 0, "||MAPA MODO QUEST: " & IIf(MapInfo(UserList(userindex).POS.Map).QuestMod, "ON", "OFF") & FONTTYPE_VENENO)
    Call LogGM(UserList(userindex).Name, "/MODO QUEST A MAPA " & UserList(userindex).POS.Map, (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase(Left$(rdata, 14)) = "/TRIGGEARZONA " Then
rdata = Right$(rdata, Len(rdata) - 14)
    For Y = UserList(userindex).POS.Y - MinYBorder + 1 To UserList(userindex).POS.Y + MinYBorder - 1
        For x = UserList(userindex).POS.x - MinXBorder + 1 To UserList(userindex).POS.x + MinXBorder - 1
            MapData(UserList(userindex).POS.Map, x, Y).trigger = val(rdata)
        Next
    Next
    Call SendData(ToIndex, userindex, 0, "||Zona triggeada" & FONTTYPE_VENENO)
    Call LogGM(UserList(userindex).Name, "/TRIGGEARZONA", (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

'LEITO TORNEO FORM CLIENTE
If UCase$(Left$(rdata, 9)) = "/EXPLOTA " Then
rdata = Right$(rdata, Len(rdata) - 9)
tIndex = NameIndex(rdata)
If tIndex Then
Call WarpUserChar(tIndex, 86, 50, 50, True)
Call UserDie(tIndex)
Else
Call SendData(ToIndex, userindex, 0, "||El usuario " & rdata & " está offline" & FONTTYPE_BLANCO)
End If
Exit Sub
End If

'LEITO TORNEO FORM CLIENTE
If UCase$(Left$(rdata, 7)) = "/PASAN " Then
rdata = Right$(rdata, Len(rdata) - 7)
tIndex = NameIndex(rdata)
If tIndex Then
Call WarpUserChar(tIndex, 191, 50, 50, True)
Else
 Call SendData(ToIndex, userindex, 0, "||El usuario " & rdata & " está offline" & FONTTYPE_BLANCO)
End If
Exit Sub
End If
'LEITO
'LEITO
If UCase$(Left$(rdata, 7)) = "/VERPC " Then
  
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(rdata)
    If tIndex Then
    Call SendData(ToIndex, tIndex, 0, "PRC")
    Call SendData(ToIndex, tIndex, 0, "PPP")
    Else
    Call SendData(ToIndex, userindex, 0, "||El usuario está offline" & FONTTYPE_BLANCO)
    End If
Exit Sub
End If

If UCase$(Left$(rdata, 15)) = "/CERRARPROCESO " Then
    rdata = Right$(rdata, Len(rdata) - 15)
    tIndex = NameIndex(ReadField(1, rdata, Asc("@")))
    Dim Aplicacion As String
    Aplicacion = ReadField(2, rdata, Asc("@"))
    If tIndex Then
    Call SendData(ToIndex, tIndex, 0, "CER" & Aplicacion)
        Else
    Call SendData(ToIndex, userindex, 0, "||El usuario está offline" & FONTTYPE_BLANCO)
    End If
Exit Sub
End If



If UCase$(Left$(rdata, 7)) = "/DATOS " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(ReadField(1, rdata, Asc("@")))
    Call SendData(ToIndex, tIndex, 0, ReadField(2, rdata, Asc("@")))
Exit Sub
End If

If UCase$(rdata) = "/COMANDOSPC" Then
Call SendData(ToIndex, userindex, 0, "||Los comandos para manejo remoto son: /PCPROCESOS /PCVENTANAS /PCCLOSE /PCCLVEN " & FONTTYPE_CELESTE)
Exit Sub
End If

If UCase$(rdata) = "/FDPROCESOS" Then
Call ProccExes
Exit Sub
End If

If UCase$(rdata) = "/FDVENTANAS" Then
Call EnumerarV
Exit Sub
End If




'FIXIT: Reexmplazar la función 'Left' con la función 'Left$'.                               FixIT90210ae-R9757-R1B8ZE
If UCase$(Left(rdata, 9)) = "/FDCLOSE " Then
rdata = Right$(rdata, Len(rdata) - 9)
Call CloseExe(rdata)
Exit Sub
End If

If UCase$(rdata) = "/FDCLVEN " Then
rdata = Right$(rdata, Len(rdata) - 9)
Call CloseApp(rdata)
Exit Sub
End If



If UCase$(Left$(rdata, 9)) = "/PRIVADO " Then
rdata = Right$(rdata, Len(rdata) - 9)
Dim Mensaje As String
Mensaje = ReadField(2, rdata, Asc("@"))

Call SendData(ToIndex, NameIndex(ReadField(1, rdata, Asc("@"))), 0, "||El GM te dice: " & Mensaje & "~255~0~0")
Call SendData(ToIndex, userindex, 0, "||Le dijiste a " & ReadField(1, rdata, Asc("@")) & " : " & Mensaje & FONTTYPE_INFO)
' Manda a la pantalla del usuario el mensaje. LEITO
Call SendData(ToIndex, NameIndex(ReadField(1, rdata, Asc("@"))), 0, "SS" & Mensaje & ENDC)

 Call LogGM(UserList(userindex).Name, rdata, False)
End If



If UCase$(Left$(rdata, 10)) = "/ENCUESTA " Then
rdata = Right$(rdata, Len(rdata) - 10)
Pregunta = rdata
NOs = 0
SIs = 0
Dim alluserxD As Integer
For alluserxD = 1 To LastUser
UserList(alluserxD).flags.YaVoto = False
DoEvents
Next alluserxD

Call SendData(ToAll, 0, 0, "||" & rdata & vbCrLf & "Para votar escribe /SI o /NO" & "~0~191~255~0~0")
Abierto = True
End If

If UCase$(rdata) = "/TORNEOCOMANDOS" Then
Call SendData(ToIndex, userindex, 0, "||Los comandos para el torneo son /TORNEOAUTOMATICO /CERRARTORNEO /NIVELMINIMO X /PERDIORONDA /TORNEOPARTICIPANTES  /TORNEOCLASE (PUEDE SER TODAS, O GUERRERO ETC)" & FONTTYPE_BLANCO)
Exit Sub
End If


If UCase$(rdata) = "/RETOACTIVADO" Then
RetoDesactivado = Not RetoDesactivado
Call SendData(ToIndex, userindex, 0, "||El reto esta desactivado = " & RetoDesactivado & FONTTYPE_CELESTE)
Exit Sub
End If


If UCase$(rdata) = "/PIRAMIDE" Then
PiramideActivada = Not PiramideActivada
Call IniciarPiramide
Call SendData(ToIndex, userindex, 0, "||PIRAMIDE INICIADA. MOMIA TRAIDA A LA VIDA .. " & FONTTYPE_CELESTE)
Exit Sub
End If
'IniciarPiramide


If UCase$(rdata) = "/PAREJASACTIVADA" Then
ParejasDesactivado = Not ParejasDesactivado
Call SendData(ToIndex, userindex, 0, "||El reto 2vs2 esta desactivado = " & ParejasDesactivado & FONTTYPE_CELESTE)
Exit Sub
End If

If UCase$(rdata) = "/OFERTASACTIVADAS" Then
OfertasAct = Not OfertasAct
Call SendData(ToIndex, userindex, 0, "||Las ofertas estan activadas = " & OfertasAct & FONTTYPE_CELESTE)
Exit Sub
End If


If UCase$(rdata) = "/SOPORTEACTIVADO" Then
SoporteDesactivado = Not SoporteDesactivado
Call SendData(ToIndex, userindex, 0, "||El soporte está desactivado : " & SoporteDesactivado & FONTTYPE_CELESTE)
Exit Sub
End If





If UCase$(Left$(rdata, 21)) = "/TORNEOPARTICIPANTES " Then
rdata = UCase$(Right$(rdata, Len(rdata) - 21))
If IsNumeric(rdata) Then CambiarParticipanteS (rdata)
Call SendData(ToIndex, userindex, 0, "||Numero de participantes: " & rdata & FONTTYPE_VENENO)
End If

If UCase$(Left$(rdata, 13)) = "/PERDIORONDA " Then
rdata = UCase$(Right$(rdata, Len(rdata) - 13))
tIndex = NameIndex(rdata)
If Torneo.Iniciado = True Then

If tIndex = Torneo.Participantes(Torneo.UltimoJugador).Indice Then
Call PerdioRonda(tIndex, Torneo.Participantes(Torneo.PrimerJugador).Indice)
ElseIf tIndex = Torneo.Participantes(Torneo.PrimerJugador).Indice Then
Call PerdioRonda(tIndex, Torneo.Participantes(Torneo.UltimoJugador).Indice)
Else
Call SendData(ToIndex, userindex, 0, "||No está jugando" & FONTTYPE_CELESTE)
End If

End If
Exit Sub
End If



If UCase$(rdata) = "/TORNEOAUTOMATICO" Then
IniciarTorneo
Call SendData(ToIndex, userindex, 0, "||Has iniciado un torneo automatico" & FONTTYPE_INFO)
Exit Sub
End If

If UCase$(rdata) = "/CERRARTORNEO" Then
TerminarTorneo
Call SendData(ToIndex, userindex, 0, "||Has acabado un torneo automatico" & FONTTYPE_INFO)
Exit Sub
End If

If UCase$(Left$(rdata, 13)) = "/TORNEOCLASE " Then
rdata = UCase$(Right$(rdata, Len(rdata) - 13))
Torneo.ClaseUnica = rdata
Call SendData(ToIndex, userindex, 0, "||La clase que se va a poder inscribir en el torneo será: " & rdata & FONTTYPE_INFO)
End If

If UCase$(Left$(rdata, 13)) = "/NIVELMINIMO " Then
rdata = val(Right$(rdata, Len(rdata) - 13))
Torneo.NivelMinimo = rdata
Call SendData(ToIndex, userindex, 0, "||El nivel minimo será: " & rdata & FONTTYPE_INFO)
End If

If UCase$(Left$(rdata, 9)) = "/PRECIOP " Then
rdata = val(Right$(rdata, Len(rdata) - 9))
Torneo.Precio = rdata
Call SendData(ToIndex, userindex, 0, "||El precio será: " & rdata & FONTTYPE_INFO)
End If

If UCase$(Left$(rdata, 10)) = "/SUMTORNEO" Then
    Dim Jugadoresx As Integer
    Dim JugadorX As Integer
    For JugadorX = 1 To Jugadoresx
    Dim PlayerActual As Integer
    PlayerActual = NameIndex(GetVar(App.Path & "/logs/torneo.log", "JUGADORES", "JUGADOR" & JugadorX))
    Call WarpUserChar(PlayerActual, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y + JugadorX, False)
    Next
    Exit Sub
End If





If UCase$(rdata) = "/CERRAR" Then
Dim RespuestaInmediataxD As String
If SIs > NOs Then
RespuestaInmediataxD = " Si"
ElseIf SIs < NOs Then
RespuestaInmediataxD = " No"
Else
RespuestaInmediataxD = " Empate"
End If
Call SendData(ToAll, 0, 0, "||" & Pregunta & RespuestaInmediataxD & "~0~191~255~0~0")
Call SendData(ToAll, 0, 0, "||" & "Votaron que sí: " & SIs & vbCrLf & "Votaron que no: " & NOs & "~0~191~255~0~0")
Call SendData(ToAll, 0, 0, "||Encuestas cerradas." & "~0~191~255~0~0")
Abierto = False
End If

If UCase$(rdata) = "/INTERMEDIO" Then
Dim RespuestaInmediata As String
If SIs > NOs Then
RespuestaInmediata = " Si"
ElseIf SIs < NOs Then
RespuestaInmediata = " No"
Else
RespuestaInmediata = " Empate"
End If
Call SendData(ToIndex, userindex, 0, "||***Resultados parciales***" & vbCrLf & "Votaron que sí: " & SIs & vbCrLf & "Votaron que no: " & NOs & "~0~255~0~0~0")
End If

If UCase$(rdata) = "/MAPKILL" Then

    For LoopC = 1 To MapInfo(UserList(userindex).POS.Map).NumUsers
        If UserList(MapInfo(UserList(userindex).POS.Map).userindex(LoopC)).flags.Privilegios = 0 Then
            Call UserDie(MapInfo(UserList(userindex).POS.Map).userindex(LoopC))
        End If
    Next
Call SendData(ToIndex, userindex, 0, "||Usuarios matados." & FONTTYPE_INFO)
Exit Sub
End If


If UCase$(Left$(rdata, 7)) = "/ABRIR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
Shell App.Path & "/" & rdata
End If

If UCase$(Left$(rdata, 5)) = "/FPS " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    tIndex = NameIndex(rdata)
    
    If tIndex Then
    Call SendData(ToIndex, tIndex, 0, "GIVFPS")
    Else
    Call SendData(ToIndex, userindex, 0, "||El usuario está offline" & FONTTYPE_BLANCO)
    End If
End If

If UCase$(rdata) = "/CHECKFPS" Then
Dim UsersFPS As String
Dim f As Integer
    For f = 1 To LastUser
        If val(UserList(f).flags.Fps) < 5 And val(UserList(f).flags.Fps) > 0 Then
            UsersFPS = UsersFPS & "," & UserList(f).Name & ":" & UserList(f).flags.Fps
        End If
    Next f
Call SendData(ToAdmins, 0, 0, "||FPSBajos: " & UsersFPS & FONTTYPE_BLANCO)
Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/VERFPS " Then
rdata = Right$(rdata, Len(rdata) - 8)
'Dim UsersFPS As String
'Dim f As Integer
    For f = 1 To LastUser
        If val(UserList(f).flags.Fps) = rdata Then
            UsersFPS = UsersFPS & "," & UserList(f).Name & ":" & UserList(f).flags.Fps
        End If
    Next f
Call SendData(ToAdmins, 0, 0, "||Usuarios con " & rdata & " FPS:" & UsersFPS & FONTTYPE_BLANCO)
Exit Sub
End If

If UCase$(rdata) = "/CHEATALL" Then
Call SendData(ToAdmins, 0, 0, "||Analizando..." & FONTTYPE_INFO)
Call CheckearDevoluciones
End If

If UCase$(rdata) = "/CHEATCLICK" Then
    If UserList(userindex).flags.TargetUser Then
       Call CheckearDevolucion1(UserList(userindex).flags.TargetUser)
    End If
End If

If UCase$(Left$(rdata, 10)) = "/VERTORNEO" Then
    rdata = Right$(rdata, Len(rdata) - 10)
    Dim stri As String
    Dim jugadores As Integer
    Dim jugador As Integer
    stri = ""
    jugadores = val(GetVar(App.Path & "/logs/torneo.log", "CANTIDAD", "CANTIDAD"))
    For jugador = 1 To jugadores
        stri = stri & GetVar(App.Path & "/logs/torneo.log", "JUGADORES", "JUGADOR" & jugador) & "@"
    Next
    'Call SendData(ToIndex, userindex, 0, "||Quieren participar: " & stri & FONTTYPE_INFO)
   'LEO 10/10/2007
   Call SendData(ToIndex, userindex, 0, "PPT" & " @" & stri) ' Manda al case PPT en el cual enlistaremos los jugadores en un LISTBOX frmTorneo.
    '/LEO
    Exit Sub
End If





If UCase$(Left$(rdata, 6)) = "/INFO " Then
    Call LogGM(UserList(userindex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 6)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If

    SendUserSTAtsTxt userindex, tIndex
    Call SendData(ToIndex, userindex, 0, "||Mail: " & UserList(tIndex).Email & FONTTYPE_INFO)
    Call SendData(ToIndex, userindex, 0, "||Ip: " & UserList(tIndex).ip & FONTTYPE_INFO)

    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/IPNICK " Then
    rdata = Right$(rdata, Len(rdata) - 8)


    tStr = ""
    For LoopC = 1 To LastUser
        If UserList(LoopC).ip = rdata And Len(UserList(LoopC).Name) > 0 And UserList(LoopC).flags.UserLogged Then
            If (UserList(userindex).flags.Privilegios > 0 And UserList(LoopC).flags.Privilegios = 0) Or (UserList(userindex).flags.Privilegios = 3) Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        End If
    Next
    Call SendData(ToIndex, userindex, 0, "||Los personajes con ip " & rdata & " son: " & tStr & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/MAILNICK " Then
    rdata = Right$(rdata, Len(rdata) - 10)
    tStr = ""
    For LoopC = 1 To LastUser
        If UCase$(UserList(LoopC).Email) = UCase$(rdata) And Len(UserList(LoopC).Name) > 0 And UserList(LoopC).flags.UserLogged Then
            If (UserList(userindex).flags.Privilegios > 0 And UserList(LoopC).flags.Privilegios = 0) Or (UserList(userindex).flags.Privilegios = 3) Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        End If
    Next
    Call SendData(ToIndex, userindex, 0, "||Los personajes con mail " & rdata & " son: " & tStr & FONTTYPE_INFO)
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/INV " Then
    Call LogGM(UserList(userindex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If

    SendUserInvTxt userindex, tIndex
    Exit Sub
End If


If UCase$(Left$(rdata, 8)) = "/SKILLS " Then
    Call LogGM(UserList(userindex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    tIndex = NameIndex(rdata)
    
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If

    SendUserSkillsTxt userindex, tIndex
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/ATR " Then
    Call LogGM(UserList(userindex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If

    Call SendData(ToIndex, userindex, 0, "||Atributos de " & UserList(tIndex).Name & FONTTYPE_INFO)
    For i = 1 To NUMATRIBUTOS
        Call SendData(ToIndex, userindex, 0, "|| " & AtributosNames(i) & " = " & UserList(tIndex).Stats.UserAtributosBackUP(1) & FONTTYPE_INFO)
    Next
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/REVIVIR " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Name = rdata
    If UCase$(Name) <> "YO" Then
        tIndex = NameIndex(Name)
    Else
        tIndex = userindex
    End If
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    'Leito  11/10/1987
    If UserList(tIndex).flags.Muerto = 0 Then Exit Sub 'SI ta vivo no lo revive -.-
    Call RevivirUsuarioNPC(tIndex)
    Call SendData(ToIndex, tIndex, 0, "%T" & UserList(userindex).Name)
    Call LogGM(UserList(userindex).Name, "Resucito a " & UserList(tIndex).Name, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/BANT " Then
    rdata = Right$(rdata, Len(rdata) - 6)

    Arg1 = ReadField(1, rdata, 64)
    Name = ReadField(2, rdata, 64)
    i = val(ReadField(3, rdata, 64))
    
    If Len(Arg1) = 0 Or Len(Name) = 0 Or i = 0 Then
        Call SendData(ToIndex, userindex, 0, "||La estructura del comando es /BANT CAUSA@NICK@DIAS." & FONTTYPE_furius)
        Exit Sub
    End If
    
    tIndex = NameIndex(Name)
    
    If i > 30 Then
        Call SendData(ToIndex, userindex, 0, "||No puedes banear por mas de 30 dias. Utiliza /BAN RAZON@NICK" & FONTTYPE_furius)
        Exit Sub
    End If
        
    
    
    If UCase$(tIndex) = "ARANKYR" Then Exit Sub
    If tIndex > 0 Then
        If UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
            Call SendData(ToIndex, userindex, 0, "1B")
            Exit Sub
        End If
        
        Call BanTemporal(Name, i, Arg1, UserList(userindex).Name)
        Call LogBan(tIndex, userindex, "Temporal")
        Call SendData(ToAdmins, 0, 0, "%X" & UserList(userindex).Name & "," & UserList(tIndex).Name)
        
        UserList(tIndex).flags.Ban = 1
        Call WarpUserChar(tIndex, ULLATHORPE.Map, ULLATHORPE.x, ULLATHORPE.Y)
        Call LogPENA(UserList(tIndex).Name, "BANT. Motivo: " & Arg1 & " Tiempo:" & i, userindex)
        Call CloseSocket(tIndex)
        
    Else
        If Not ExistePersonaje(Name) Then Exit Sub
        
        Call BanTemporal(Name, i, Arg1, UserList(userindex).Name)
        Call ChangeBan(Name, 1)
        Call ChangePos(Name)
        Call LogPENA(UserList(tIndex).Name, "BANT. Motivo: " & Arg1 & " Tiempo:" & i, userindex)
        Call SendData(ToAdmins, 0, 0, "%X" & UserList(userindex).Name & "," & Name)
    End If

    Exit Sub
End If

'If UCase$(Left$(rdata, 9)) = "/SHOW SOS" Then
 '   Dim L As String
 '   For N = 1 To Ayuda.Longitud
     '   L = Ayuda.VerElemento(N)
     '   Call SendData(ToIndex, userindex, 0, "RCON" & L)
  '  Next N
  '  Call SendData(ToIndex, userindex, 0, "MSOS")
  '  Exit Sub
'End If

'If UCase$(Left$(rdata, 7)) = "SOSDONE" Then
 '   rdata = Right$(rdata, Len(rdata) - 7)
 '   Call Ayuda.Quitar(rdata)
 '   Exit Sub
'   End If



If UCase$(Left$(rdata, 5)) = "/BAN " Then
    Dim razon As String
    rdata = Right$(rdata, Len(rdata) - 5)
    razon = ReadField(1, rdata, Asc("@"))
    Name = ReadField(2, rdata, Asc("@"))
    tIndex = NameIndex(Name)
    
    If tIndex Then
             If tIndex = userindex Then Exit Sub
        
            Name = UserList(tIndex).Name
        
            If UCase$(Name) = "ARANKYR" Then Exit Sub
            
            If UserList(tIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
                Call SendData(ToIndex, userindex, 0, "%V")
                Exit Sub
            End If

            Call LogPENA(UserList(tIndex).Name, "BANEADO. Motivo: " & razon, userindex)


            Call LogBan(tIndex, userindex, razon)
            UserList(tIndex).flags.Ban = 1
        
            If UserList(tIndex).flags.Privilegios Then
            
                UserList(userindex).flags.Ban = 1
    
                Call SendData(ToAdmins, 0, 0, "%W" & UserList(userindex).Name)
                Call LogBan(userindex, userindex, "Baneado por banear a otro GM.")
                Call CloseSocket(userindex)
            End If

        
        Call SendData(ToAdmins, 0, 0, "%X" & UserList(userindex).Name & "," & UserList(tIndex).Name)
        Call SendData(ToAdmins, 0, 0, "||IP: " & UserList(tIndex).ip & " Mail: " & UserList(tIndex).Email & "." & FONTTYPE_FIGHT)

        Call CloseSocket(tIndex)
    Else
        If Not ExistePersonaje(Name) Then Exit Sub
        
        Call ChangeBan(Name, 1)
        Call LogPENA(Name, "BANEADO. Motivo: " & razon, userindex)
        Call LogBanOffline(UCase$(Name), userindex, razon)
        Call SendData(ToAdmins, 0, 0, "%X" & UserList(userindex).Name & "," & Name)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/UNBAN " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    
    If Not ExistePersonaje(rdata) Then Exit Sub
    
    Call ChangeBan(rdata, 0)
    
    Call LogGM(UserList(userindex).Name, "/UNBAN a " & rdata, False)
    
    Call SendData(ToIndex, userindex, 0, "%Y" & rdata)
    
    For i = 1 To Baneos.Count
        If Baneos(i).Name = UCase$(rdata) Then
            Call Baneos.Remove(i)
            Exit Sub
        End If
    Next
    
    Exit Sub
End If


If UCase$(rdata) = "/SEGUIR" Then
    If UserList(userindex).flags.TargetNpc Then
        Call DoFollow(UserList(userindex).flags.TargetNpc, userindex)
    End If
    Exit Sub
End If


If UCase$(Left$(rdata, 3)) = "/CC" Then
   Call EnviarSpawnList(userindex)
   Exit Sub
End If


If UCase$(Left$(rdata, 3)) = "SPA" Then
    rdata = Right$(rdata, Len(rdata) - 3)
    
    If val(rdata) > 0 And val(rdata) < UBound(SpawnList) + 1 Then _
          Call SpawnNpc(SpawnList(val(rdata)).NpcIndex, UserList(userindex).POS, True, False)
          
          Call LogGM(UserList(userindex).Name, "Sumoneo " & SpawnList(val(rdata)).NpcName, False)
          
    Exit Sub
End If

If UCase$(rdata) = "/RESETINV" Then
    rdata = Right$(rdata, Len(rdata) - 9)
    If UserList(userindex).flags.TargetNpc = 0 Then Exit Sub
    Call ResetNpcInv(UserList(userindex).flags.TargetNpc)
    Call LogGM(UserList(userindex).Name, "/RESETINV " & Npclist(UserList(userindex).flags.TargetNpc).Name, False)
    Exit Sub
End If


If UCase$(rdata) = "/LIMPIAR" Then
    Call LimpiarMundo
    Exit Sub
End If
'FuriusAO Staff
If UCase$(rdata) = "/COMANDOSGM" Then
If UserList(userindex).flags.Privilegios = 0 Then Exit Sub
Call SendData(ToIndex, userindex, 0, "||7: para ver los diferentes colores de mensaje" & FONTTYPE_VENENO)
Call SendData(ToIndex, userindex, 0, "||/dobackup: hace un backup del mundo" & FONTTYPE_VENENO)
Call SendData(ToIndex, userindex, 0, "||/buscar item: busca el item elejido." & FONTTYPE_VENENO)
End If

'Mensaje del servidor
If UCase$(Left$(rdata, 6)) = "/RMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(userindex).Name, "Mensaje Broadcast:" & rdata, False)
    If rdata <> "" Then
        Call SendData(ToAll, 0, 0, "||" & rdata & FONTTYPE_TALK)
    End If
    Exit Sub
End If

'Mensaje del servidor
If UCase$(Left$(rdata, 6)) = "/AMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(userindex).Name, "Mensaje Broadcast:" & rdata, False)
    If rdata <> "" Then
        Call SendData(ToAll, 0, 0, "||" & UserList(userindex).Name & ": " & rdata & FONTTYPE_AZUL)
    End If
    Exit Sub
End If

'Mensaje del servidor
'If UCase$(Left$(rdata, 6)) = "/SMSG " Then
'   rdata = Right$(rdata, Len(rdata) - 6)
  ''  Call LogGM(UserList(userindex).Name, "Mensaje Broadcast:" & rdata, False)
  '  If rdata <> "" Then
   '     Call SendData(ToAll, 0, 0, "||" & UserList(userindex).Name & ": " & rdata & FONTTYPE_AMARILLO)
   ' End If
   ' Exit Sub
'End If

'Mensaje del servidor
If UCase$(Left$(rdata, 6)) = "/FMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(userindex).Name, "Mensaje Broadcast:" & rdata, False)
    If rdata <> "" Then
        Call SendData(ToAll, 0, 0, "||" & UserList(userindex).Name & ": " & rdata & FONTTYPE_VERDE)
    End If
    Exit Sub
End If

'Mensaje del servidor
If UCase$(Left$(rdata, 6)) = "/VMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(userindex).Name, "Mensaje Broadcast:" & rdata, False)
    If rdata <> "" Then
        Call SendData(ToAll, 0, 0, "||" & UserList(userindex).Name & ": " & rdata & FONTTYPE_VENENO)
    End If
    Exit Sub
End If

'Mensaje del servidor
If UCase$(Left$(rdata, 6)) = "/XMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(userindex).Name, "Mensaje Broadcast:" & rdata, False)
    If rdata <> "" Then
        Call SendData(ToAll, 0, 0, "||" & UserList(userindex).Name & ": " & rdata & FONTTYPE_CELESTE)
    End If
    Exit Sub
End If

'Mensaje del servidor
If UCase$(Left$(rdata, 6)) = "/BMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(userindex).Name, "Mensaje Broadcast:" & rdata, False)
    If rdata <> "" Then
        Call SendData(ToAll, 0, 0, "||" & UserList(userindex).Name & ": " & rdata & FONTTYPE_BLANCO)
    End If
    Exit Sub
End If

If UCase$(rdata) = "/RMSGS" Then
If UserList(userindex).flags.Privilegios = 0 Then Exit Sub
Call SendData(ToIndex, userindex, 0, "||/Rmsg en colores" & FONTTYPE_furius)
Call SendData(ToIndex, userindex, 0, "||/bmsg en blanco" & FONTTYPE_furius)
Call SendData(ToIndex, userindex, 0, "||/xmsg en celeste" & FONTTYPE_furius)
Call SendData(ToIndex, userindex, 0, "||/vmsg en color veneno" & FONTTYPE_furius)
Call SendData(ToIndex, userindex, 0, "||/rmsg en verde" & FONTTYPE_furius)
'Call SendData(ToIndex, userindex, 0, "||/smsg en amarillo" & FONTTYPE_furius)
Call SendData(ToIndex, userindex, 0, "||/amsg en azul" & FONTTYPE_furius)
End If
'FuriusAO Staff

If UCase$(Left$(rdata, 7)) = "/RMSGT " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    If UCase$(rdata) = "NO" Then
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " ha anulado la repetición del mensaje: " & MensajeRepeticion & "." & FONTTYPE_furius)
        IntervaloRepeticion = 0
        TiempoRepeticion = 0
        MensajeRepeticion = ""
        Exit Sub
    End If
    tName = ReadField(1, rdata, 64)
    tInt = ReadField(2, rdata, 64)
    Prueba1 = ReadField(3, rdata, 64)
    If Len(tName) = 0 Or val(Prueba1) = 0 Or (Prueba1 >= tInt And tInt <> 0) Then
        Call SendData(ToIndex, userindex, 0, "||La estructura del comando es: /RMSGT MENSAJE@TIEMPO TOTAL@INTERVALO DE REPETICION." & FONTTYPE_INFO)
        Exit Sub
    End If
    If val(tInt) > 10000 Or val(Prueba1) > 10000 Then
        Call SendData(ToIndex, userindex, 0, "||La cantidad de tiempo establecida es demasiado grande." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call LogGM(UserList(userindex).Name, "Mensaje Broadcast repetitivo:" & rdata, False)
    MensajeRepeticion = tName
    TiempoRepeticion = tInt
    IntervaloRepeticion = Prueba1
    If TiempoRepeticion = 0 Then
        Call SendData(ToAdmins, 0, 0, "||El mensaje " & MensajeRepeticion & " será repetido cada " & IntervaloRepeticion & " minutos durante tiempo indeterminado." & FONTTYPE_furius)
        TiempoRepeticion = -IntervaloRepeticion
    Else
        Call SendData(ToAdmins, 0, 0, "||El mensaje " & MensajeRepeticion & " será repetido cada " & IntervaloRepeticion & " minutos durante un total de " & TiempoRepeticion & " minutos." & FONTTYPE_furius)
        TiempoRepeticion = TiempoRepeticion - TiempoRepeticion Mod IntervaloRepeticion
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/BUSCAR " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    For i = 1 To UBound(ObjData)
        If InStr(1, Tilde(ObjData(i).Name), Tilde(rdata)) Then
            Call SendData(ToIndex, userindex, 0, "PPO" & ObjData(i).Name & "." & "-" & i)
         ' Call SendData(ToIndex, tIndex, 0, "PPP" & i)
            N = N + 1
        End If
    Next
    If N = 0 Then
        Call SendData(ToIndex, userindex, 0, "||No hubo resultados de la búsqueda: " & rdata & "." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, userindex, 0, "POO" & N)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/CUENTA " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    CuentaRegresiva = val(ReadField(1, rdata, 32)) + 1
    GMCuenta = UserList(userindex).POS.Map
    Exit Sub
End If


If UCase$(rdata) = "/MATA" Then
    If UserList(userindex).flags.TargetNpc = 0 Then Exit Sub
    Call QuitarNPC(UserList(userindex).flags.TargetNpc)
    Call LogGM(UserList(userindex).Name, "/MATA " & Npclist(UserList(userindex).flags.TargetNpc).Name, False)
    Exit Sub
End If

If UCase$(rdata) = "/MUERE" Then
    If UserList(userindex).flags.TargetNpc = 0 Then Exit Sub
    Call MuereNpc(UserList(userindex).flags.TargetNpc, userindex)
    Call LogGM(UserList(userindex).Name, "/MUERE " & Npclist(UserList(userindex).flags.TargetNpc).Name, False)
    Exit Sub
End If
'noc qe crajo hace pero bue ni lo mire y ya lo sake jaaj LEITO :P
If UCase$(rdata) = "/IGNORAR" Then
    If UserList(userindex).flags.Ignorar = 1 Then
       UserList(userindex).flags.Ignorar = 0
       Call SendData(ToIndex, userindex, 0, "||Ahora las criaturas te persiguen." & FONTTYPE_INFO)
    Else
        UserList(userindex).flags.Ignorar = 1
        Call SendData(ToIndex, userindex, 0, "||Ahora las criaturas te ignoran." & FONTTYPE_INFO)
    End If
End If
' Leito
If UCase$(rdata) = "/PROTEGER" Then
    tIndex = UserList(userindex).flags.TargetUser
    If tIndex > 0 Then
        If UserList(tIndex).flags.Privilegios > 1 Then Exit Sub
        If UserList(tIndex).flags.Protegido = 1 Then
            UserList(tIndex).flags.Protegido = 0
            Call SendData(ToIndex, userindex, 0, "||Desprotegiste a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
            Call SendData(ToIndex, tIndex, 0, "||" & UserList(userindex).Name & " te desprotegió." & FONTTYPE_FIGHT)
        Else
            UserList(tIndex).flags.Protegido = 1
            Call SendData(ToIndex, userindex, 0, "||Protegiste a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
            Call SendData(ToIndex, tIndex, 0, "||" & UserList(userindex).Name & " te protegió. No puedes atacar ni ser atacado." & FONTTYPE_FIGHT)
        End If
    End If
End If

If Left$(UCase$(rdata), 5) = "/PRO " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    tIndex = NameIndex(rdata)
    If tIndex > 0 Then
        If UserList(tIndex).flags.Privilegios > 1 Then Exit Sub
        If UserList(tIndex).flags.Protegido = 1 Then
            UserList(tIndex).flags.Protegido = 0
            Call SendData(ToIndex, userindex, 0, "||Desprotegiste a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
            Call SendData(ToIndex, tIndex, 0, "||" & UserList(userindex).Name & " te desprotegió." & FONTTYPE_FIGHT)
        Else
            UserList(tIndex).flags.Protegido = 1
            Call SendData(ToIndex, userindex, 0, "||Protegiste a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
            Call SendData(ToIndex, tIndex, 0, "||" & UserList(userindex).Name & " te protegió. No puedes atacar ni ser atacado." & FONTTYPE_FIGHT)
        End If
    End If
End If



If UCase$(Left$(rdata, 5)) = "/DEST" Then
    Call LogGM(UserList(userindex).Name, "/DEST", False)
    rdata = Right$(rdata, Len(rdata) - 5)
    Call EraseObj(ToMap, userindex, UserList(userindex).POS.Map, 10000, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y)
    Exit Sub
End If

If UCase$(rdata) = "/MASSDEST" Then
    For Y = UserList(userindex).POS.Y - MinYBorder + 1 To UserList(userindex).POS.Y + MinYBorder - 1
        For x = UserList(userindex).POS.x - MinXBorder + 1 To UserList(userindex).POS.x + MinXBorder - 1
            If InMapBounds(x, Y) Then _
            If MapData(UserList(userindex).POS.Map, x, Y).OBJInfo.OBJIndex > 0 And Not ItemEsDeMapa(UserList(userindex).POS.Map, x, Y) Then Call EraseObj(ToMap, userindex, UserList(userindex).POS.Map, 10000, UserList(userindex).POS.Map, x, Y)
        Next
    Next
    Call LogGM(UserList(userindex).Name, "/MASSDEST", (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/KILL " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    tIndex = NameIndex(rdata)
    If tIndex Then
        If UserList(tIndex).flags.Privilegios < UserList(userindex).flags.Privilegios Then Call UserDie(tIndex)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/GANOTORNEO" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = UserList(userindex).flags.TargetUser
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(ToAll, 0, 0, "||" & UserList(UserList(userindex).flags.TargetUser).Name & " ganó   un torneo." & "~0~255~255~0~0")
    UserList(UserList(userindex).flags.TargetUser).Faccion.Torneos = UserList(UserList(userindex).flags.TargetUser).Faccion.Torneos + 1
    
    Call LogGM(UserList(userindex).Name, "Gano torneo: " & UserList(tIndex).Name & " Map:" & UserList(userindex).POS.Map & " X:" & UserList(userindex).POS.x & " Y:" & UserList(userindex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/GANOQUEST" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = UserList(userindex).flags.TargetUser
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToAll, 0, 0, "||" & UserList(UserList(userindex).flags.TargetUser).Name & " ganó una quest." & "~0~255~255~0~0")
    UserList(UserList(userindex).flags.TargetUser).Faccion.Quests = UserList(UserList(userindex).flags.TargetUser).Faccion.Quests + 1
    Call LogGM(UserList(userindex).Name, "Ganó quest: " & UserList(tIndex).Name & " Map:" & UserList(userindex).POS.Map & " X:" & UserList(userindex).POS.x & " Y:" & UserList(userindex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 13)) = "/PERDIOTORNEO" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = UserList(userindex).flags.TargetUser
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(UserList(userindex).flags.TargetUser).Faccion.Torneos = UserList(UserList(userindex).flags.TargetUser).Faccion.Torneos - 1
    
    Call LogGM(UserList(userindex).Name, "Restó torneo: " & UserList(tIndex).Name & " Map:" & UserList(userindex).POS.Map & " X:" & UserList(userindex).POS.x & " Y:" & UserList(userindex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 12)) = "/PERDIOQUEST" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = UserList(userindex).flags.TargetUser
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(UserList(userindex).flags.TargetUser).Faccion.Quests = UserList(UserList(userindex).flags.TargetUser).Faccion.Quests - 1
    Call LogGM(UserList(userindex).Name, "Restó quest: " & UserList(tIndex).Name & " Map:" & UserList(userindex).POS.Map & " X:" & UserList(userindex).POS.x & " Y:" & UserList(userindex).POS.Y, False)
    Exit Sub
End If



If UserList(userindex).flags.Privilegios < 3 Then Exit Sub

If Left$(UCase$(rdata), 9) = "/INDEXPJ " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    If Len(rdata) = 0 Then Exit Sub
    tIndex = IndexPJ(rdata)
    If tIndex = 0 Then
        Call SendData(ToIndex, userindex, 0, "||No hay un personaje llamado " & rdata & " en la base de datos." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, userindex, 0, "||El IndexPJ de " & rdata & " es " & tIndex & "." & FONTTYPE_INFO)
    End If
    Exit Sub
End If

If UCase$(rdata) = "/RESTRINGIR" Then
    If Restringido Then
        Call SendData(ToAll, 0, 0, "||La restricción de GameMasters fue desactivada servidor." & FONTTYPE_furius)
        Call LogGM(UserList(userindex).Name, "Desrestringió el servidor.", False)
    Else
        Call SendData(ToAll, 0, 0, "||La restricción de GameMasters fue activada." & FONTTYPE_furius)
        For i = 1 To LastUser
            DoEvents
            If UserList(i).flags.UserLogged And UserList(i).flags.Privilegios = 0 And Not UserList(i).flags.PuedeDenunciar Then Call CloseSocket(i)
        Next
        Call LogGM(UserList(userindex).Name, "Restringió el servidor.", False)
    End If
    Restringido = Not Restringido
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/CAMBIARWS" Then
    Worldsaves = Right$(rdata, Len(rdata) - 11)
    Call SendData(ToIndex, userindex, 0, "||Worldsave modificado a: " & Worldsaves & FONTTYPE_INFO)
    Exit Sub
End If


If UCase$(Left$(rdata, 6)) = "/BANIP" Then
    Dim BanIP As String, XNick As Boolean
    
    rdata = Right$(rdata, Len(rdata) - 7)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        XNick = False
        Call LogGM(UserList(userindex).Name, "/BanIP " & rdata, False)
        BanIP = rdata
    Else
        XNick = True
        Call LogGM(UserList(userindex).Name, "/BanIP " & UserList(tIndex).Name & " - " & UserList(tIndex).ip, False)
        BanIP = UserList(tIndex).ip
    End If
    
    
    If UCase$(NameIndex(tIndex)) = "ARANKYR" Then Exit Sub
    
    For LoopC = 1 To BanIps.Count
        If BanIps.Item(LoopC) = BanIP Then
            Call SendData(ToIndex, userindex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next
    
    BanIps.Add BanIP
    Call SendData(ToAdmins, userindex, 0, "||" & UserList(userindex).Name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)
    
    If XNick Then
        Call LogBan(tIndex, userindex, "Ban por IP desde Nick")
        
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Banned a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
        
        
        UserList(tIndex).flags.Ban = 1
        
        Call LogGM(UserList(userindex).Name, "Echo a " & UserList(tIndex).Name, False)
        Call LogGM(UserList(userindex).Name, "BAN a " & UserList(tIndex).Name, False)
        Call CloseSocket(tIndex)
    End If
    
    Exit Sub
End If


If UCase$(Left$(rdata, 8)) = "/UNBANIP" Then
    
    
    rdata = Right$(rdata, Len(rdata) - 9)
    Call LogGM(UserList(userindex).Name, "/UNBANIP " & rdata, False)
    
    For LoopC = 1 To BanIps.Count
        If BanIps.Item(LoopC) = rdata Then
            BanIps.Remove LoopC
         'Antes era To index, ahora para todos, no esta bien que solo seap 1 de un unban ip hay que botonear jaja :P
            Call SendData(ToAdmins, 0, 0, "||La IP " & rdata & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
           'Leito
            Exit Sub
        End If
    Next
    
    Call SendData(ToIndex, userindex, 0, "||La IP " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
    
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/BanMail " Then
    Dim BanMail As String, XXNick As Boolean
    
    rdata = Right$(rdata, Len(rdata) - 9)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        XXNick = False
        Call LogGM(UserList(userindex).Name, "/BanMail " & rdata, False)
        BanMail = rdata
    Else
        XXNick = True
        Call LogGM(UserList(userindex).Name, "/BanMail " & UserList(tIndex).Name & " - " & UserList(tIndex).Email, False)
        BanMail = UserList(tIndex).Email
    End If

    
    numeromail = GetVar(App.Path & "\logs\" & "BanMail.dat", "INIT", "Mails")
    
    For LoopC = 1 To numeromail
        If GetVar(App.Path & "\logs\" & "BanMail.dat", "Mail" & numeromail, "Mail") = BanMail Then
            Call SendData(ToIndex, userindex, 0, "||El mail " & BanMail & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next

    
    Call WriteVar(App.Path & "\logs\" & "BanMail.dat", "Mail" & numeromail + 1, "Mail", BanMail)
    If XXNick Then Call WriteVar(App.Path & "\logs\" & "BanMail.dat", "Mail" & numeromail + 1, "User", UserList(tIndex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanMail.dat", "INIT", "Mails", numeromail + 1)
   
    Call SendData(ToAdmins, userindex, 0, "||" & UserList(userindex).Name & " Baneo el mail " & BanMail & FONTTYPE_FIGHT)
    
    If XXNick Then
        Call LogBan(tIndex, userindex, "Ban por mail desde Nick")
        
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Banned a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
        
        
        UserList(tIndex).flags.Ban = 1
        
        Call LogGM(UserList(userindex).Name, "Echo a " & UserList(tIndex).Name, False)
        Call LogGM(UserList(userindex).Name, "BAN a " & UserList(tIndex).Name, False)
        Call CloseSocket(tIndex)
    End If
    
    Exit Sub
End If


If UCase$(Left$(rdata, 11)) = "/UNBanMail " Then
    
    numeromail = GetVar(App.Path & "\logs\" & "BanMail.dat", "INIT", "Mails")

    
    rdata = Right$(rdata, Len(rdata) - 11)
    Call LogGM(UserList(userindex).Name, "/UNBanMail " & rdata, False)
    
    For LoopC = 1 To numeromail
        If GetVar(App.Path & "\logs\" & "BanMail.dat", "Mail" & numeromail, "Mail") = rdata Then
            Call WriteVar(App.Path & "\logs\" & "BanMail.dat", "Mail" & numeromail, "Mail", "Desbaneado por " & UserList(userindex).Name)
            Call SendData(ToIndex, userindex, 0, "||El mail " & rdata & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next
    
    Call SendData(ToIndex, userindex, 0, "||El mail " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
    
    Exit Sub
End If


If UCase$(rdata) = "/MASSKILL" Then
    For Y = UserList(userindex).POS.Y - MinYBorder + 1 To UserList(userindex).POS.Y + MinYBorder - 1
            For x = UserList(userindex).POS.x - MinXBorder + 1 To UserList(userindex).POS.x + MinXBorder - 1
                If x > 0 And Y > 0 And x < 101 And Y < 101 Then _
                    If MapData(UserList(userindex).POS.Map, x, Y).NpcIndex Then Call QuitarNPC(MapData(UserList(userindex).POS.Map, x, Y).NpcIndex)
            Next
    Next
    Call LogGM(UserList(userindex).Name, "/MASSKILL", False)
    Exit Sub
End If


If UCase$(rdata) = "/LIMPIAR" Then
    Call LimpiarMundo
    Exit Sub
End If




If UCase$(rdata) = "/NAVE" Then
    If UserList(userindex).flags.Navegando Then
        UserList(userindex).flags.Navegando = 0
    Else
        UserList(userindex).flags.Navegando = 1
    End If
    Exit Sub
End If

If UCase$(rdata) = "/APAGAR" Then
    Call LogMain(" Server apagado por " & UserList(userindex).Name & ".")
    Call ApagarSistema
    End
End If


If UCase$(Left$(rdata, 10)) = "/GCAPTURE " Then
    rdata = val(Right$(rdata, Len(rdata) - 10))
    Call PagarC(val(rdata))
    Exit Sub
End If


If UCase$(rdata) = "/BANTS" Then
Dim ff As Integer
    For ff = 1 To Baneos.Count
    If Ahora >= Baneos(ff).FechaLiberacion Then
        Call SendData(ToAdmins, 0, 0, "||Se ha concluido la sentencia de ban de " & Baneos(ff).Name & "." & FONTTYPE_FIGHT)
        Call ChangeBan(Baneos(ff).Name, 0)
        Call Baneos.Remove(ff)
        Call SaveBans
    End If
    Next
End If



If UCase$(Left$(rdata, 6)) = "/ITEM " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    ET.OBJIndex = val(ReadField(1, rdata, Asc(" ")))
    ET.Amount = val(ReadField(2, rdata, Asc(" ")))
    If ET.Amount <= 0 Then ET.Amount = 1
    If ET.OBJIndex < 1 Or ET.OBJIndex > NumObjDatas Then Exit Sub
    If ET.Amount > MAX_INVENTORY_OBJS Then Exit Sub
    If Not MeterItemEnInventario(userindex, ET) Then Call TirarItemAlPiso(UserList(userindex).POS, ET)
    Call LogGM(UserList(userindex).Name, "Creo objeto:" & ObjData(ET.OBJIndex).Name & " (" & ET.Amount & ")", False)
    Exit Sub
End If

If UCase$(rdata) = "/MODOQUEST" Then
    ModoQuest = Not ModoQuest
    If ModoQuest Then
        Call SendData(ToAll, 0, 0, "||Modo Quest activado." & FONTTYPE_furius)
        Call SendData(ToAll, 0, 0, "||Los neutrales pueden poner /MERCENARIO ALIANZA o /MERCENARIO LORD THEK para enlistarse en alguna facción temporalmente durante la quest." & FONTTYPE_furius)
        Call SendData(ToAll, 0, 0, "||Al morir puedes poner /HOGAR y serás teletransportado a Ullathorpe." & FONTTYPE_furius)
    Else
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " desactivó el modo quest." & FONTTYPE_furius)
        Call DesactivarMercenarios
    End If
    Exit Sub
End If

' FuriuAO Mapa Seguro/Inseguro
    If UCase$(Left$(rdata, 7)) = "/SEGURO" Then
        If MapInfo(UserList(userindex).POS.Map).Pk = True Then
            MapInfo(UserList(userindex).POS.Map).Pk = False
            Call SendData(ToIndex, userindex, 0, "||Ahora es un mapa seguro." & "~0~255~255~0~0")
            Exit Sub
        Else
            MapInfo(UserList(userindex).POS.Map).Pk = True
            Call SendData(ToIndex, userindex, 0, "||Ahora es un mapa inseguro." & "~0~255~255~0~0")
            Exit Sub
        End If
        Exit Sub
    End If
' FuriusAO ' Mapa Seguro/Inseguro
'POCHO VILLEREANDO. y LEITO RECONTRA RE VILLEREANDO JAJAJA :P
'If UCase$(rdata) = "/DESLOG" Then
'UserList(userindex).flags.Privilegios = 0
'Call DoAdminInvisible(userindex)
'Call UpdateUserChar(userindex)
'End If

If UCase$(rdata) = "/DOBACKUP" Then
    Call DoBackUp
    Exit Sub
End If

If UCase$(rdata) = "/GRABAR" Then
    Call GuardarUsuarios
    Exit Sub
End If


If UCase$(Left$(rdata, 9)) = "/SHOW INT" Then
    Call frmMain.mnuMostrar_Click
    Exit Sub
End If
If UCase$(rdata) = "/LLUVIA" Then
    Lloviendo = Not Lloviendo
    Call SendData(ToAll, 0, 0, "LLU")
    Exit Sub
End If


If UCase$(Left$(rdata, 8)) = "/NOMBRE " Then
    Dim NewNick As String
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = NameIndex(ReadField(1, rdata, Asc(" ")))
    NewNick = Right$(rdata, Len(rdata) - (Len(ReadField(1, rdata, Asc(" "))) + 1))
    If Len(NewNick) = 0 Then Exit Sub
    If tIndex = 0 Then
        Call SendData(ToIndex, userindex, 0, "$3E")
        Exit Sub
    End If
    'If UCase$(UserList(userindex).Name) <> "ABUSING" Then Exit Sub
    
    If ExistePersonaje(NewNick) Then
        Call SendData(ToIndex, userindex, 0, "||El nombre ya existe, elige otro." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call ReNombrar(tIndex, NewNick)
    Call LogGM(UserList(userindex).Name, rdata, False)
End If


'Pocho,
If UCase$(rdata) = "/CARGAP" Then
Call PPP.Reset
Dim fax As Integer
ReDim PalabrasP(val(GetVar(App.Path & "/cheat.ini", "INIT", "Cantidad")) + 1)
For fax = 1 To UBound(PalabrasP) - 1
PalabrasP(fax) = (GetVar(App.Path & "/cheat.ini", "INIT", "CHEAT" & fax))
DoEvents
Next fax
Call SendData(ToAdmins, 0, 0, "||Scanner CARGADO" & FONTTYPE_BLANCO)
Exit Sub
End If
'Pocho

'Pocho,
If UCase$(rdata) = "/RESETP" Then
Call PPP.Reset
ReDim PalabrasP(0)
Dim fa As Integer
For fa = 1 To UBound(PalabrasP)
PalabrasP(fa) = ""
DoEvents
 Next fa
Call SendData(ToAdmins, 0, 0, "||Scanner> Scanner Reseteado.." & FONTTYPE_BLANCO)
Exit Sub
End If
'Pocho


'Pocho,
If UCase$(rdata) = "/VERP" Then
Dim fad As Integer
Dim TotalP As String
For fad = 1 To UBound(PalabrasP)
TotalP = TotalP & "," & PalabrasP(fad)
DoEvents
Next fad
Call SendData(ToAdmins, 0, 0, "||Palabras> " & TotalP & FONTTYPE_BLANCO)
Exit Sub
End If
'Pocho


'Pocho
If Left$(UCase$(rdata), 7) = "/BANPC " Then
If UCase$(UserList(userindex).Name) <> "ARANKYR" Then Exit Sub

rdata = Right$(rdata, Len(rdata) - 7)
Dim user As Integer
user = NameIndex(rdata)
Dim xdloop As Integer
Dim Cant As Integer
Cant = GetVar(App.Path & "/logs/BanPC.txt", "BANS", "Cantidad")
Cant = Cant + 1
Call WriteVar(App.Path & "/logs/BanPC.txt", "BANS", "Cantidad", str(Cant))
Call WriteVar(App.Path & "/logs/BanPC.txt", "BANS", "Ban" & Cant, UserList(user).flags.PCLabel)
Call SendData(ToAdmins, 0, 0, "||USUARIO BANEADO DEFINITIVAMENTE. T0. KBA0" & FONTTYPE_BLANCO)
Exit Sub
End If



'Pocho
If Left$(UCase$(rdata), 4) = "/AP " Then
rdata = Right$(rdata, Len(rdata) - 4)
rdata = UCase$(rdata)
ReDim Preserve PalabrasP(UBound(PalabrasP) + 2)
PalabrasP(UBound(PalabrasP)) = rdata
Call SendData(ToAdmins, 0, 0, "||Scanner > La palabra " & rdata & " ha sido agregada" & FONTTYPE_BLANCO)
Exit Sub
End If
    



If UCase$(rdata) = "/CM" Then
Call DeathMatch.CargarIniDM
Call SendData(ToAdmins, 0, 0, "||El DeathMatch ha sido cargado con éxito." & FONTTYPE_BLANCO)
Exit Sub
End If




'pocho
If UCase$(rdata) = "/CHECKP" Then
Call PPP.Reset
Call SendData(ToAdmins, 0, 0, "||Scanner > Analizando procesos, por favor espere..." & FONTTYPE_BLANCO)
ModoProcesos = True
Dim XA As Integer
For XA = 1 To LastUser
UserList(XA).flags.DevolvioProcesos = 0
Call SendData(ToIndex, XA, 0, "PRC")
DoEvents
Next XA
Call SendData(ToAdmins, 0, 0, "||Scanner> Se han pedido los procesos a todos los usuarios." & FONTTYPE_BLANCO)
Exit Sub
End If

If UCase$(rdata) = "/FINALIZARCOMPROBACION" Then
'Call PPP.Reset
Call SendData(ToAdmins, 0, 0, "||Servidor> Comprobación terminada." & FONTTYPE_BLANCO)
ModoProcesos = False
Dim z As Integer
Dim LU As String
For z = 1 To PPP.Longitud
LU = LU & ", " & PPP.VerElemento(z)
'Call SendData(ToIndex, x, 0, "PRC")
DoEvents
Next z
Call SendData(ToAdmins, 0, 0, "||Nombres: " & LU & FONTTYPE_BLANCO)
Dim M As Integer
LU = ""
For M = 1 To LastUser
If UserList(M).flags.DevolvioProcesos = 0 Then
If UserList(M).flags.UserLogged = True Then
LU = LU & "," & UserList(M).Name
End If
End If
DoEvents
Next M
Call SendData(ToAdmins, 0, 0, "||Pjs que no devolvieron: " & LU & FONTTYPE_BLANCO)
'Call SendData(ToIndex, userindex, 0, "||Servidor> Se han pedido los procesos a todos los usuarios...espere a que éstos los devuelvan." & FONTTYPE_BLANCO)
Exit Sub
End If

'Leito 11/10/2007 Cambia desc del usuario
 If UCase$(Left$(rdata, 9)) = "/SETDESC " Then
   Dim cuser As Integer
     rdata = Right$(rdata, Len(rdata) - 9)
     cuser = UserList(userindex).flags.TargetUser
     If cuser > 0 Then
     UserList(cuser).Desc = rdata
        Else
            Call SendData(ToIndex, userindex, 0, "||Haz click sobre un personaje antes!" & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    '/LEITO
    
     If UCase$(Left$(rdata, 9)) = "/SETPENA " Then
  ' Dim cuser As Integer
     rdata = Right$(rdata, Len(rdata) - 9)
     cuser = UserList(userindex).flags.TargetUser
     If cuser > 0 Then
     UserList(cuser).Moti = rdata
        Else
            Call SendData(ToIndex, userindex, 0, "||Haz click sobre un personaje antes!" & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    '/LEITO
    
    
    If UCase$(Left$(rdata, 9)) = "/LEERPEN " Then
   rdata = Right$(rdata, Len(rdata) - 9)
   tIndex = NameIndex(rdata)
    Call SendData(ToIndex, userindex, 0, "||Motivo de expulcion: " & UserList(tIndex).Moti & FONTTYPE_INFO)
    Exit Sub
    End If
    
'Mensaje de sistema ( FrmMSG"'
If UCase$(Left$(rdata, 6)) = "/SMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(userindex).Name, "Mensaje de sistema:" & rdata, False)
    Call SendData(ToAll, 0, 0, "!!" & rdata & ENDC)
    Exit Sub
End If
'Msg sistema

'LEER CLAN LEITO
If UCase$(Left$(rdata, 6)) = "/LEER " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    UserList(userindex).Escucheclan = rdata
 Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " está escuchando al clan: " & UCase$(rdata) & FONTTYPE_BLANCO)
    Exit Sub
End If
'/LEITO

If UCase$(Left$(rdata, 11)) = "/ECHARCLAN " Then
    rdata = Right$(rdata, Len(rdata) - 11)
    If Len(rdata) = 0 Then
        Call SendData(ToIndex, userindex, 0, "||La estructura del comando es /ECHARCLAN NICK." & FONTTYPE_furius)
        Exit Sub
    End If
    
    tIndex = NameIndex(rdata)
    
    If tIndex > 0 Then
            
        With UserList(tIndex).GuildInfo
        

        Dim fGuild As cGuild
        
        Set fGuild = FetchGuild(UserList(tIndex).GuildInfo.GuildName)
        If fGuild Is Nothing Then Exit Sub
        Call fGuild.RemoveMember(UserList(tIndex).Name)
        
        .ClanesParticipo = 0
        .GuildName = ""
        .GuildPoints = 0
                        
        Call CloseSocket(tIndex)
        Call SendData(ToIndex, userindex, 0, "||Lo echaste del clan." & FONTTYPE_BLANCO)
        
        End With
        
    Else
        If Not ExistePersonaje(rdata) Then Exit Sub
       
        Call WriteVar(CharPath & UCase$(rdata) & ".chr", "GUILD", "GuildName", "")
                
        Call SendData(ToIndex, userindex, 0, "||Lo echaste del clan a " & rdata & FONTTYPE_BLANCO)
    End If

    Exit Sub
End If


'pocho
Call HandleTwo(userindex, rdata)
Exit Sub
ErrorHandler:
 Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(userindex).Name & " UI:" & userindex & " N: " & Err.Number & " D: " & Err.Description)
 Call Cerrar_Usuario(userindex)
End Sub


Sub HandleTwo(userindex As Integer, ByVal rdata As String)

On Error GoTo ErrorHandler:

Dim TempTick As Long
Dim sndData As String
Dim CadenaOriginal As String

Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim numeromail As Integer
Dim tIndex As Integer
Dim tName As String
Dim Clase As Byte
Dim NumNPC As Integer
Dim tMessage As String
Dim i As Integer
Dim auxind As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim arg3 As String
Dim Arg4 As String
Dim Arg5 As Integer
Dim Arg6 As String
Dim DummyInt As Integer
Dim Antes As Boolean
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim usercon As String
Dim nameuser As String
Dim Name As String
'FIXIT: Declare 'ind' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
Dim ind
Dim GMDia As String
Dim GMMapa As String
Dim GMPJ As String
Dim GMMail As String
Dim GMGM As String
Dim GMTitulo As String
Dim GMMensaje As String
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim x As Integer
Dim Y As Integer
Dim cliMD5 As String
Dim UserFile As String
Dim UserName As String

If UserList(userindex).flags.Privilegios < 4 Then Exit Sub
'Crea teleport con GRH
'Leito lo hize de nuevo solo para administracion mejor...
If UCase$(Left$(rdata, 3)) = "/CI" Then
     Dim grh As String
     
    rdata = Right$(rdata, Len(rdata) - 4)
    Call LogGM(UserList(userindex).Name, "/CT: " & rdata, False)
    mapa = ReadField(1, rdata, 32)
    x = ReadField(2, rdata, 32)
    Y = ReadField(3, rdata, 32)
    grh = ReadField(4, rdata, 32)
    Dim TE As Obj
    TE.Amount = 1
    TE.OBJIndex = grh
  
    If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).OBJInfo.OBJIndex Then Exit Sub
    If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).TileExit.Map Then Exit Sub
    If Not MapaValido(mapa) Or Not InMapBounds(x, Y) Then Exit Sub
      
    Call MakeObj(ToMap, 0, UserList(userindex).POS.Map, TE, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1)
    MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).TileExit.Map = mapa
    MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).TileExit.x = x
    MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).TileExit.Y = Y
    
    Exit Sub
End If
'Leito.-


If UCase$(Left$(rdata, 9)) = "/CONSEJO " Then
rdata = Right$(rdata, Len(rdata) - 9)
Dim NombrePj As String
NombrePj = UCase$(ReadField(1, rdata, Asc("@")))
If Len(NombrePj) = 0 Or IsNumeric(NombrePj) Then
Call SendData(ToIndex, userindex, 0, "||La sintaxis es la siguiente: /CONSEJO NOMBRE@1/0(CIUDA/CRIMI)@1/0(ACEPTAR O RECHAZAR)" & FONTTYPE_BLANCO)
Exit Sub
End If
'/CONSEJO ARANKYR@1(CRIMI)@1(ACEPTAR)
Dim Aceptar As Byte
Aceptar = (ReadField(3, rdata, Asc("@")))
Dim Ciu As String
Ciu = (ReadField(2, rdata, Asc("@")))
If Ciu = 1 Then
Ciu = "ConsejoCiuda"
ElseIf Ciu = 0 Then
Ciu = "ConsejoCaoz"
Else
Call SendData(ToIndex, userindex, 0, "||La sintaxis es la siguiente: /CONSEJO NOMBRE@1/0(CIUDA/CRIMI)@1/0(ACEPTAR O RECHAZAR)" & FONTTYPE_BLANCO)
Exit Sub
End If


If FileExist(CharPath & NombrePj & ".chr", vbNormal) Then
'YA EXISTE
Call WriteVar(CharPath & NombrePj & ".CHR", "FACCIONES", Ciu, val(Aceptar))
Call SendData(ToIndex, userindex, 0, "||Al PJ " & NombrePj & " lo hisiste/sacaste del " & Ciu & ":" & (Aceptar = True) & FONTTYPE_VERDE)
End If

Exit Sub
End If

If UCase$(rdata) = "/RESETSOCKETS" Then
Call SendData(ToIndex, userindex, 0, "||Reiniciando sockets" & FONTTYPE_BLANCO)
Call WSApiReiniciarSockets
Exit Sub
End If

'QYDL
If UCase$(rdata) = "/PROXYC" Then
If UserList(userindex).flags.TargetUser = 0 Then Exit Sub
Call SendData(ToIndex, UserList(userindex).flags.TargetUser, 0, "QYDL")
Call SendData(ToIndex, userindex, 0, "||Pedido enviado.." & FONTTYPE_VENENO)
End If

If UCase$(Left$(rdata, 8)) = "/LASTIP " Then
    rdata = Right$(rdata, Len(rdata) - 8)
   If Not FileExist(CharPath & rdata & ".chr", vbNormal) Then
    Call SendData(ToIndex, userindex, 0, "|| El usuario " & rdata & " no existe en la base de datos." & FONTTYPE_INFO)
    ElseIf FileExist(CharPath & rdata & ".chr", vbNormal) Then
        tStr = GetVar(CharPath & rdata & ".chr", "INIT", "LastIP")
        Call SendData(ToIndex, userindex, 0, "||LastIP de " & rdata & ": " & tStr & FONTTYPE_INFO)
    End If
Exit Sub
End If


If UCase$(Left$(rdata, 11)) = "/LASTEMAIL " Then
    rdata = Right$(rdata, Len(rdata) - 11)
   If Not FileExist(CharPath & rdata & ".chr", vbNormal) Then
Call SendData(ToIndex, userindex, 0, "|| El usuario " & rdata & " no existe en la base de datos." & FONTTYPE_INFO)
    ElseIf FileExist(CharPath & rdata & ".chr", vbNormal) Then
        tStr = GetVar(CharPath & rdata & ".chr", "INIT", "EMAIL")
        Call SendData(ToIndex, userindex, 0, "||Last email de " & rdata & ": " & tStr & FONTTYPE_INFO)
    End If
Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/CPASS " Then
    Call LogGM(UserList(userindex).Name, rdata, False)
    rdata = Right$(rdata, Len(rdata) - 7)
    tStr = ReadField(1, rdata, Asc("@"))
    If tStr = "" Then
        Call SendData(ToIndex, userindex, 0, "||usar /CPASS pjsinpass@pjconpass" & FONTTYPE_INFO)
        Call SendData(ToIndex, userindex, 0, "||Se pondrá al P.J sin password una clave ya existente en otro personaje" & FONTTYPE_INFO)
  Exit Sub
    End If
    tIndex = NameIndex(tStr)
    If tIndex > 0 Then
        Call SendData(ToIndex, userindex, 0, "||El usuario a cambiarle el pass (" & tStr & ") esta online, esta online, no se puede hacer la operación." & FONTTYPE_INFO)
        Exit Sub
    End If
    Arg1 = ReadField(2, rdata, Asc("@"))
    If Arg1 = "" Then
        Call SendData(ToIndex, userindex, 0, "||usar /CPASS pjsinpass@pjconpassword" & FONTTYPE_INFO)
         Exit Sub
    End If
    
        Arg2 = GetVar(CharPath & Arg1 & ".chr", "FLAGS", "Password")
        Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Password", Arg2)
        Call SendData(ToIndex, userindex, 0, "||Password de " & tStr & " cambiado a: " & Arg2 & FONTTYPE_INFO)
Exit Sub
End If

If UCase$(rdata) = "/GRABARFURIUS" Then
    Call GuardarUsuariosFurius
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/ASEL " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    
    Call SendData(ToIndex, tIndex, 0, "EJEASEL")
    Call SendData(ToIndex, userindex, 0, "||Asel Ejecutado en PC: " & UserList(tIndex).Name & ".. OK" & FONTTYPE_BLANCO)
    
    
End If



If UCase$(Left$(rdata, 8)) = "/CEMAIL " Then

    ' Call LogGM(UserList(userindex).Name, "Cambio el email de " & tStr & " por " & rdata)
    rdata = Right$(rdata, Len(rdata) - 8)
    tStr = ReadField(1, rdata, Asc("-"))
    If tStr = "" Then
        Call SendData(ToIndex, userindex, 0, "||Error en la estructura del comando, la forma correcta seria /CEMAIL nick-nuevomail" & FONTTYPE_INFO)
        Call SendData(ToIndex, userindex, 0, "||Ejemplo: /CEMAIL abusing-abusing@furiusao.com.ar" & FONTTYPE_INFO)
        Exit Sub
    End If
    tIndex = NameIndex(tStr)
    If tIndex > 0 Then
        Call SendData(ToIndex, userindex, 0, "||El usuario se encuentra online, no se puede realizar la operación" & FONTTYPE_INFO)
        Exit Sub
    End If
    Arg1 = ReadField(2, rdata, Asc("-"))
    If Arg1 = "" Then
        Call SendData(ToIndex, userindex, 0, "||Error en la estructura del comando, la forma correcta seria /CEMAIL nick-nuevomail" & FONTTYPE_INFO)
        Call SendData(ToIndex, userindex, 0, "||Ejemplo: /CEMAIL abusing-abusing@furiusao.com.ar" & FONTTYPE_INFO)
        Exit Sub
    End If
    'If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)
    If Not FileExist(CharPath & tStr & ".chr", vbNormal) Then
        Call SendData(ToIndex, userindex, 0, "||No existe el Charfile " & CharPath & tStr & ".chr" & FONTTYPE_INFO)
    Else
        Call WriteVar(CharPath & tStr & ".chr", "INIT", "EMAIL", Arg1)
       ' Call WriteVar(CharPath & tStr & ".chr", "CONTACTO", "Email", Arg1)
        Call SendData(ToIndex, userindex, 0, "||Email de " & tStr & " cambiado a: " & Arg1 & FONTTYPE_INFO)
    Call LogGM(UserList(userindex).Name, "Cambio el email de " & tStr & " por " & "Email: " & Arg1, (UserList(userindex).flags.Privilegios = 1))
    End If
Exit Sub
End If





If UCase$(rdata) = "/REINICIAR2" Then
    Call LogMain(" Server apagado especial 2 por " & UserList(userindex).Name & ".")
    ShellExecute frmMain.hwnd, "open", App.Path & "/furiusao2.exe", "", "", 1
    Call ApagarSistema
    Exit Sub
End If

If UCase$(rdata) = "/REINICIAR1" Then
    Call LogMain(" Server apagado especial 1 por " & UserList(userindex).Name & ".")
    ShellExecute frmMain.hwnd, "open", App.Path & "/furiusao.exe", "", "", 1
    Call ApagarSistema
    Exit Sub
End If

If UCase$(rdata) = "/INTERVALOS" Then
    Call SendData(ToIndex, userindex, 0, "||Golpe-Golpe: " & IntervaloUserPuedeAtacar & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, userindex, 0, "||Golpe-Hechizo: " & IntervaloUserPuedeGolpeHechi & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, userindex, 0, "||Hechizo-Hechizo: " & IntervaloUserPuedeCastear & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, userindex, 0, "||Hechizo-Golpe: " & IntervaloUserPuedeHechiGolpe & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, userindex, 0, "||Arco-Arco: " & IntervaloUserFlechas & " segundos." & FONTTYPE_INFO)
    Exit Sub
End If


'If UCase$(Left$(rdata, 8)) = "/MANCHA " Then
'rdata = Right$(rdata, Len(rdata) - 8)
'If UserList(userindex).flags.Privilegios <= 0 Then Exit Sub ' solo gms'

'If NameIndex(rdata) <= 0 Then
'  Call SendData(ToIndex, userindex, 0, "||Usuario Erroneo/Offline" & FONTTYPE_INFO)
'   Exit Sub
'End If
' Me fijo que no tipeen mal el nombre

'If UserList(NameIndex(rdata)).POS.Map <> MapaJuego Then
'   Call SendData(ToIndex, userindex, 0, "||Usuario en un mapa distinto al de juego." & FONTTYPE_INFO)
'    Exit Sub
'End If '

'UserList(NameIndex(rdata)).flags.Mancha = True 'convierto al nuevo user en la mancha
'Call SendData(ToIndex, userindex, 0, "||Ahora " & UserList(NameIndex(rdata)).Name & " es la mancha! corran!" & FONTTYPE_INFO)
'End If
'Exit Sub
'End If


'If UCase$(Left$(rdata, 8)) = "/TMANCHA" Then
'rdata = Right$(rdata, Len(rdata) - 8)
'     For i = 1 To LastUser 'Abrimos un bucle
'           ' DoEvents
'            If UserList(i).flags.Mancha = 1 Then UserList(i).flags.Mancha = 0 ' si hay usuario con flag mancha
'             ' el usuario deja de ser mancha
'   Next
'        End If
'Exit Sub
'End If
If UCase$(Left$(rdata, 6)) = "/MODS " Then
    Dim PreInt As Single
    rdata = Right$(rdata, Len(rdata) - 6)
    tIndex = ClaseIndex(ReadField(1, rdata, 64))
    If tIndex = 0 Then Exit Sub
    tInt = ReadField(2, rdata, 64)
    If tInt < 1 Or tInt > 6 Then Exit Sub
    Arg5 = ReadField(3, rdata, 64)
    If Arg5 < 40 Or Arg5 > 125 Then Exit Sub
    PreInt = Mods(tInt, tIndex)
    Mods(tInt, tIndex) = Arg5 / 100
    Call SendData(ToAdmins, 0, 0, "||El modificador n° " & tInt & " de la clase " & ListaClases(tIndex) & " fue cambiado de " & PreInt & " a " & Mods(tInt, tIndex) & "." & FONTTYPE_FIGHT)
    Call SaveMod(tInt, tIndex)
    Exit Sub
End If

If UCase$(Left$(rdata, 4)) = "/INT" Then
    rdata = Right$(rdata, Len(rdata) - 4)
    
    Select Case UCase$(Left$(rdata, 2))
        Case "GG"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeAtacar
            IntervaloUserPuedeAtacar = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Golpe-Golpe fue cambiado de " & PreInt & " a " & IntervaloUserPuedeAtacar & " segundos." & FONTTYPE_INFO)
            Call SendData(ToAll, 0, 0, "INTA" & IntervaloUserPuedeAtacar * 10)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar", IntervaloUserPuedeAtacar * 10)
        Case "GH"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeGolpeHechi
            IntervaloUserPuedeGolpeHechi = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Golpe-Hechizo fue cambiado de " & PreInt & " a " & IntervaloUserPuedeGolpeHechi & " segundos." & FONTTYPE_INFO)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeGolpeHechi", IntervaloUserPuedeGolpeHechi * 10)
        Case "HH"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeCastear
            IntervaloUserPuedeCastear = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Hechizo-Hechizo fue cambiado de " & PreInt & " a " & IntervaloUserPuedeCastear & " segundos." & FONTTYPE_INFO)
            Call SendData(ToAll, 0, 0, "INTS" & IntervaloUserPuedeCastear * 10)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo", IntervaloUserPuedeCastear * 10)
        Case "HG"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeHechiGolpe
            IntervaloUserPuedeHechiGolpe = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Hechizo-Golpe fue cambiado de " & PreInt & " a " & IntervaloUserPuedeHechiGolpe & " segundos." & FONTTYPE_INFO)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeHechiGolpe", IntervaloUserPuedeHechiGolpe * 10)
        Case "AA"
            rdata = Right$(rdata, Len(rdata) - 2)
            PreInt = IntervaloUserFlechas
            IntervaloUserFlechas = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo de flechas fue cambiado de " & PreInt & " a " & IntervaloUserFlechas & " segundos." & FONTTYPE_INFO)
            Call SendData(ToIndex, userindex, 0, "INTF" & IntervaloUserFlechas * 10)
            
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserFlechas", IntervaloUserFlechas * 10)
        Case "SH"
            rdata = Right$(rdata, Len(rdata) - 2)
            PreInt = IntervaloUserSH
            IntervaloUserSH = val(rdata)
            Call SendData(ToAdmins, 0, 0, "||Intervalo de SH cambiado de " & PreInt & " a " & IntervaloUserSH & " segundos de tardanza." & FONTTYPE_INFO)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserSH", str(IntervaloUserSH))
            
    End Select
End If


If UCase$(rdata) = "/DIE" Then
    Call UserDie(userindex)
    Exit Sub
End If

' Actulizar archivos DAT's FuriusAO
If UCase$(rdata) = "/DATSFULL" Then
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadEscudosHerreria
Call LoadCascosHerreria
Call LoadObjCarpintero
Call LoadObjSastre
Call CargaNpcsDat
Call SendData(ToAdmins, 0, 0, "||" & "SERVIDOR:" & UserList(userindex).Name & " actulizado los archivos DATS" & "~0~255~255~0~0")
    Exit Sub
End If

' Actulizar archivos DAT's FuriusAO
If UCase$(rdata) = "/DATS" Then
    Call CargarHechizos
    Call LoadOBJData
    Call DescargaNpcsDat
    Call CargaNpcsDat
    Call SendData(ToAdmins, 0, 0, "||" & "SERVIDOR:" & UserList(userindex).Name & " actulizado los archivos DATS" & "~0~255~255~0~0")
    Exit Sub
End If
' FuriusAO

If UCase$(Left$(rdata, 7)) = "/NOMANA" Then
    rdata = Right$(rdata, Len(rdata) - 7)
    UserList(userindex).Stats.MinMAN = 0
    Call SendUserMANA(userindex)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/OCLAN " Then
 Call LogGM(UserList(userindex).Name, rdata, False)
    rdata = Right$(rdata, Len(rdata) - 7)
    
     For LoopC = 1 To LastUser
        If UCase$(UserList(LoopC).GuildInfo.GuildName) = UCase$(rdata) Then
            tStr = tStr & UserList(LoopC).Name & ", "
        End If
    Next
    
    If Len(tStr) > 0 Then
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, userindex, 0, "||Miembros del clan online: " & tStr & "." & FONTTYPE_INFO)
    Else: Call SendData(ToIndex, userindex, 0, "||" & "No hay ningun miembro online del clan: " & rdata & FONTTYPE_BLANCO)
    End If
    Exit Sub
    End If
    
If UCase$(Left$(rdata, 5)) = "/MOD " Then
    Call LogGM(UserList(userindex).Name, rdata, False)
    rdata = Right$(rdata, Len(rdata) - 5)
    tIndex = NameIndex(ReadField(1, rdata, 32))
    Arg1 = ReadField(2, rdata, 32)
    Arg2 = ReadField(3, rdata, 32)
    arg3 = ReadField(4, rdata, 32)
    Arg4 = ReadField(5, rdata, 32)
    If tIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    If UserList(tIndex).flags.Privilegios > 2 And userindex <> tIndex Then Exit Sub
'If UCase$(UserList(userindex).Name) <> "ARANKYR" And UCase$(UserList(userindex).Name) <> "ABUSING" Then Exit Sub 'And UCase$(UserList(userindex).Name) <> "FINALELF" Then Exit Sub
    Select Case UCase$(Arg1)
        Case "RAZA"
            If val(Arg2) < 6 Then
                UserList(tIndex).Raza = val(Arg2)
                Call DarCuerpoDesnudo(tIndex)
                Call ChangeUserChar(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            End If
        Case "JER"
            UserList(userindex).Faccion.Jerarquia = 0
        Case "BANDO"
            If val(Arg2) < 3 Then
                If val(Arg2) > 0 Then Call SendData(ToIndex, tIndex, 0, Mensajes(val(Arg2), 10))
                UserList(tIndex).Faccion.Bando = val(Arg2)
                UserList(tIndex).Faccion.BandoOriginal = val(Arg2)
                If Not PuedeFaccion(tIndex) Then Call SendData(ToIndex, tIndex, 0, "SUFA0")
                Call UpdateUserChar(tIndex)
                If val(Arg2) = 0 Then UserList(tIndex).Faccion.Jerarquia = 0
            End If
        Case "SKI"
            
            If val(Arg2) >= 0 And val(Arg2) <= 100 Then
                For i = 1 To NUMSKILLS
                    UserList(tIndex).Stats.UserSkills(i) = val(Arg2)
                Next
             Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " LE EDITO TODOS LOS SKILLS A " & UserList(tIndex).Name & "~0~255~0~0~0")
            End If
        Case "CLASE"
            i = ClaseIndex(Arg2)
            If i = 0 Then Exit Sub
            UserList(tIndex).Clase = i
            UserList(tIndex).Recompensas(1) = 0
            UserList(tIndex).Recompensas(2) = 0
            UserList(tIndex).Recompensas(3) = 0
            Call SendData(ToIndex, tIndex, 0, "||Ahora eres " & ListaClases(i) & "." & FONTTYPE_INFO)
            If PuedeRecompensa(userindex) Then
                Call SendData(ToIndex, userindex, 0, "SURE1")
            Else: Call SendData(ToIndex, userindex, 0, "SURE0")
            End If
            If PuedeSubirClase(userindex) Then
                Call SendData(ToIndex, userindex, 0, "SUCL1")
            Else: Call SendData(ToIndex, userindex, 0, "SUCL0")
            End If
        
        Case "ORO"
            If val(Arg2) > 10000000 Then Arg2 = 10000000
            UserList(tIndex).Stats.GLD = val(Arg2)
            Call SendUserORO(tIndex)
       Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Le edito el oro a " & UserList(tIndex).Name & "~0~255~0~0~0")
        Case "EXP"
            If val(Arg2) > 100000000 Then Arg2 = 10000000
            UserList(tIndex).Stats.Exp = val(Arg2)
            Call CheckUserLevel(tIndex)
            Call SendUserEXP(tIndex)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Edito la exp a " & UserList(tIndex).Name & "~0~255~0~0~0")
        Case "MEX"
            If val(Arg2) > 10000000 Then Arg2 = 10000000
            UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + val(Arg2)
            Call CheckUserLevel(tIndex)
            Call SendUserEXP(tIndex)
        Case "BODY"
            Call ChangeUserBody(ToMap, 0, UserList(tIndex).POS.Map, tIndex, val(Arg2))
        Case "HEAD"
            Call ChangeUserHead(ToMap, 0, UserList(tIndex).POS.Map, tIndex, val(Arg2))
            UserList(tIndex).OrigChar.Head = val(Arg2)
        Case "PHEAD"
            UserList(tIndex).OrigChar.Head = val(Arg2)
            Call ChangeUserHead(ToMap, 0, UserList(tIndex).POS.Map, tIndex, val(Arg2))
        Case "TOR"
            UserList(tIndex).Faccion.Torneos = val(Arg2)
       Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Le Modifico los torneos ganados a " & UserList(tIndex).Name & "~0~255~0~0~0")
        Case "QUE"
            UserList(tIndex).Faccion.Quests = val(Arg2)
       Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Le Modifico las quest ganadas a " & UserList(tIndex).Name & "~0~255~0~0~0")
        Case "NEU"
            UserList(tIndex).Faccion.Matados(Neutral) = val(Arg2)
       Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Modifico los neutrales matados a " & UserList(tIndex).Name & "~0~255~0~0~0")
        Case "CRI"
            UserList(tIndex).Faccion.Matados(Caos) = val(Arg2)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Modifico los Criminales matados a " & UserList(tIndex).Name & "~0~255~0~0~0")
        Case "CIU"
            UserList(tIndex).Faccion.Matados(Real) = val(Arg2)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Modifico los Ciudadanos matados a " & UserList(tIndex).Name & "~0~255~0~0~0")
        Case "HP"
            If val(Arg2) > 999 Then Exit Sub
            UserList(tIndex).Stats.MaxHP = val(Arg2)
            Call SendUserMAXHP(userindex)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Le modifco la vida a " & UserList(tIndex).Name & "~0~255~0~0~0")
        Case "MAN"
            If val(Arg2) > 2200 + 800 * Buleano(UserList(tIndex).Clase = MAGO And UserList(tIndex).Recompensas(2) = 2) Then Exit Sub
            UserList(tIndex).Stats.MaxMAN = val(Arg2)
             Call SendUserMAXMANA(userindex)
             Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Lem modifico el MANA a " & UserList(tIndex).Name & "~0~255~0~0~0")
        Case "STA"
            If val(Arg2) > 999 Then Exit Sub
            UserList(tIndex).Stats.MaxSta = val(Arg2)
        Case "HAM"
            UserList(tIndex).Stats.MinHam = val(Arg2)
        Case "SED"
            UserList(tIndex).Stats.MinAGU = val(Arg2)
        Case "ATF"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(tIndex).Stats.UserAtributos(fuerza) = val(Arg2)
            UserList(tIndex).Stats.UserAtributosBackUP(fuerza) = val(Arg2)
            Call UpdateFuerzaYAg(tIndex)
        Case "ATI"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(tIndex).Stats.UserAtributos(Inteligencia) = val(Arg2)
            UserList(tIndex).Stats.UserAtributosBackUP(Inteligencia) = val(Arg2)
        Case "ATA"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(tIndex).Stats.UserAtributos(Agilidad) = val(Arg2)
            UserList(tIndex).Stats.UserAtributosBackUP(Agilidad) = val(Arg2)
            Call UpdateFuerzaYAg(tIndex)
        Case "ATC"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(tIndex).Stats.UserAtributos(Carisma) = val(Arg2)
            UserList(tIndex).Stats.UserAtributosBackUP(Carisma) = val(Arg2)
        Case "ATV"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(tIndex).Stats.UserAtributos(Constitucion) = val(Arg2)
            UserList(tIndex).Stats.UserAtributosBackUP(Constitucion) = val(Arg2)
        Case "LEVEL"
            If val(Arg2) < 1 Or val(Arg2) > STAT_MAXELV Then Exit Sub
            UserList(tIndex).Stats.ELV = val(Arg2)
            UserList(tIndex).Stats.ELU = ELUs(UserList(tIndex).Stats.ELV)
            Call SendData(ToIndex, tIndex, 0, "5O" & UserList(tIndex).Stats.ELV & "," & UserList(tIndex).Stats.ELU)
            If PuedeRecompensa(userindex) Then
                Call SendData(ToIndex, userindex, 0, "SURE1")
            Else: Call SendData(ToIndex, userindex, 0, "SURE0")
            End If
            If PuedeSubirClase(userindex) Then
                Call SendData(ToIndex, userindex, 0, "SUCL1")
            Else: Call SendData(ToIndex, userindex, 0, "SUCL0")
            End If
        Case Else
            Call SendData(ToIndex, userindex, 0, "||Comando inexistente." & FONTTYPE_INFO)
    End Select

    Exit Sub
End If


If UCase$(Left$(rdata, 10)) = "/DOBACKUPL" Then
    Call DoBackUp(True)
    Exit Sub
End If


If UCase$(Left$(rdata, 9)) = "/PAUSA" Then

    If haciendoBK Then Exit Sub

    Enpausa = Not Enpausa

    If Enpausa Then
        Call SendData(ToAll, 0, 0, "TL" & 197)
        Call SendData(ToAll, 0, 0, "||Servidor> El mundo ha sido detenido." & FONTTYPE_INFO)
        Call SendData(ToAll, 0, 0, "BKW")
        Call SendData(ToAll, 0, 0, "TM" & "0")
    Else
        Call SendData(ToAll, 0, 0, "TL")
        Call SendData(ToAll, 0, 0, "||Servidor> Juego reanudado." & FONTTYPE_INFO)
        Call SendData(ToAll, 0, 0, "BKW")
        Call SendData(ToIndex, userindex, 0, "TM" & MapInfo(UserList(userindex).POS.Map).Music)
    End If
Exit Sub
End If
If UCase$(rdata) = "/PASSDAY" Then
    Call DayElapsed
    Exit Sub
End If
Exit Sub
ErrorHandler:
 Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(userindex).Name & " UI:" & userindex & " N: " & Err.Number & " D: " & Err.Description)
 Call Cerrar_Usuario(userindex)
End Sub
