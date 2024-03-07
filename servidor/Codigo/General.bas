Attribute VB_Name = "General"
Global ANpc As Long
Global Anpc_host As Long

Option Explicit
Public Function MapaPorUbicacion(X As Integer, Y As Integer) As Integer
Dim i As Integer

For i = 1 To NumMaps
    If MapInfo(i).LeftPunto = X And MapInfo(i).TopPunto = Y And MapInfo(i).Zona <> Dungeon Then
        MapaPorUbicacion = i
        Exit Function
    End If
Next

End Function
Public Sub WriteBIT(Variable As Byte, POS As Byte, Value As Byte)

If ReadBIT(Variable, POS) = Value Then Exit Sub

If Value = 0 Then
    Variable = Variable - 2 ^ (POS - 1)
Else: Variable = Variable + 2 ^ (POS - 1)
End If

End Sub
Public Function Valorcito(Variable As Byte, POS As Byte, Valor As Byte) As Byte

Call WriteBIT(Variable, POS, Valor)
Valorcito = Variable

End Function
Public Function ReadBIT(Variable As Byte, POS As Byte) As Byte
Dim i As Integer

ReadBIT = Variable

For i = 7 To POS Step -1
    ReadBIT = ReadBIT Mod 2 ^ i
Next

ReadBIT = ReadBIT \ 2 ^ (POS - 1)

End Function
Public Function Enemigo(ByVal Bando As Byte) As Byte

Select Case Bando
    Case Neutral
        Enemigo = 3
    Case Real
        Enemigo = Caos
    Case Caos
        Enemigo = Real
End Select

End Function
Sub DarCuerpoDesnudo(userindex As Integer)

If UserList(userindex).flags.Navegando Then
    UserList(userindex).Char.Head = 0
    UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.BarcoObjIndex).Ropaje
    Exit Sub
End If

UserList(userindex).Char.Head = UserList(userindex).OrigChar.Head

Select Case UserList(userindex).Raza
    Case HUMANO
      Select Case UserList(userindex).Genero
        Case HOMBRE
             UserList(userindex).Char.Body = 21
        Case MUJER
             UserList(userindex).Char.Body = 39
      End Select
    Case ELFO_OSCURO
      Select Case UserList(userindex).Genero
        Case HOMBRE
             UserList(userindex).Char.Body = 32
        Case MUJER
             UserList(userindex).Char.Body = 40
      End Select
    Case ENANO
      Select Case UserList(userindex).Genero
        Case HOMBRE
             UserList(userindex).Char.Body = 53
        Case MUJER
             UserList(userindex).Char.Body = 60
      End Select
    Case GNOMO
      Select Case UserList(userindex).Genero
        Case HOMBRE
             UserList(userindex).Char.Body = 53
        Case MUJER
             UserList(userindex).Char.Body = 60
      End Select
      
    Case Else
      Select Case UserList(userindex).Genero
        Case HOMBRE
             UserList(userindex).Char.Body = 21
        Case MUJER
             UserList(userindex).Char.Body = 39
      End Select
    
End Select

UserList(userindex).flags.Desnudo = 1

End Sub
Public Function PuedeDestrabarse(userindex As Integer) As Boolean
Dim i As Byte, nPos As WorldPos

If (UserList(userindex).flags.Muerto = 0) Or (Not MapInfo(UserList(userindex).POS.Map).Pk And UserList(userindex).POS.Map <> 37) Then Exit Function

For i = NORTH To WEST
    nPos = UserList(userindex).POS
    Call HeadtoPos(i, nPos)
    If InMapBounds(nPos.X, nPos.Y) Then
        If LegalPos(nPos.Map, nPos.X, nPos.Y, CBool(UserList(userindex).flags.Navegando)) Then Exit Function
    End If
Next

PuedeDestrabarse = True

End Function
Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, X As Integer, Y As Integer, B As Byte)

Call SendData(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & B)

End Sub
Sub LimpiarMundo()
On Error Resume Next
Dim i As Integer

For i = 1 To TrashCollector.Count
    Dim d As cGarbage
    Set d = TrashCollector(1)
    Call EraseObj(ToMap, 0, d.Map, 1, d.Map, d.X, d.Y)
    Call TrashCollector.Remove(1)
    Set d = Nothing
Next

End Sub
'FIXIT: Declare 'Obj' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 0 Then

Obj.AddressFamily = AF_INET
Obj.protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = Port
Obj.backlog = 5
Obj.listen

'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
End Sub
Sub EnviarSpawnList(userindex As Integer)
Dim k As Integer, SD As String
SD = "SPL" & UBound(SpawnList) & ","

For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next

Call SendData(ToIndex, userindex, 0, SD)
End Sub
Sub EstablecerRecompensas()

Recompensas(MINERO, 1, 1).SubeHP = 120

Recompensas(MAGO, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(MAGO, 1, 1).Obj(1).Amount = 1000
Recompensas(MAGO, 1, 2).Obj(1).OBJIndex = Petrificar
Recompensas(MAGO, 1, 2).Obj(1).Amount = 1
Recompensas(MAGO, 2, 1).SubeHP = 10

Recompensas(NIGROMANTE, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(NIGROMANTE, 1, 1).Obj(1).Amount = 1000
Recompensas(NIGROMANTE, 1, 2).Obj(1).OBJIndex = Petrificar
Recompensas(NIGROMANTE, 1, 2).Obj(1).Amount = 1
Recompensas(NIGROMANTE, 2, 1).SubeHP = 15
Recompensas(NIGROMANTE, 2, 2).SubeMP = 40

Recompensas(PALADIN, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(PALADIN, 1, 1).Obj(1).Amount = 1000
Recompensas(PALADIN, 1, 2).Obj(1).OBJIndex = Petrificar
Recompensas(PALADIN, 1, 2).Obj(1).Amount = 1
Recompensas(PALADIN, 2, 1).SubeHP = 5
Recompensas(PALADIN, 2, 1).SubeMP = 10
Recompensas(PALADIN, 2, 2).SubeMP = 30

Recompensas(CLERIGO, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(CLERIGO, 1, 1).Obj(1).Amount = 1000
Recompensas(CLERIGO, 1, 2).Obj(1).OBJIndex = Petrificar
Recompensas(CLERIGO, 1, 2).Obj(1).Amount = 1
Recompensas(CLERIGO, 2, 1).SubeHP = 10
Recompensas(CLERIGO, 2, 2).SubeMP = 50

Recompensas(BARDO, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(BARDO, 1, 1).Obj(1).Amount = 1000
Recompensas(BARDO, 1, 2).Obj(1).OBJIndex = Petrificar
Recompensas(BARDO, 1, 2).Obj(1).Amount = 1
Recompensas(BARDO, 2, 1).SubeHP = 10
Recompensas(BARDO, 2, 2).SubeMP = 50

Recompensas(Druida, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(Druida, 1, 1).Obj(1).Amount = 1000
Recompensas(Druida, 1, 2).Obj(1).OBJIndex = Petrificar
Recompensas(Druida, 1, 2).Obj(1).Amount = 1
Recompensas(Druida, 2, 1).SubeHP = 15
Recompensas(Druida, 2, 2).SubeMP = 40

Recompensas(ASESINO, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(ASESINO, 1, 1).Obj(1).Amount = 1000
Recompensas(ASESINO, 1, 2).Obj(1).OBJIndex = Petrificar
Recompensas(ASESINO, 1, 2).Obj(1).Amount = 1
Recompensas(ASESINO, 2, 1).SubeHP = 10
Recompensas(ASESINO, 2, 2).SubeMP = 30

Recompensas(CAZADOR, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(CAZADOR, 1, 1).Obj(1).Amount = 1000
Recompensas(CAZADOR, 1, 2).Obj(1).OBJIndex = Petrificar
Recompensas(CAZADOR, 1, 2).Obj(1).Amount = 1
Recompensas(CAZADOR, 2, 1).SubeHP = 10
Recompensas(CAZADOR, 2, 2).SubeMP = 50

Recompensas(ARQUERO, 1, 1).Obj(1).OBJIndex = Flecha
Recompensas(ARQUERO, 1, 1).Obj(1).Amount = 1500
Recompensas(ARQUERO, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(ARQUERO, 1, 2).Obj(1).Amount = 1000
Recompensas(ARQUERO, 2, 1).SubeHP = 10

Recompensas(GUERRERO, 1, 1).Obj(1).OBJIndex = PocionVerdeNoCae
Recompensas(GUERRERO, 1, 1).Obj(1).Amount = 80
Recompensas(GUERRERO, 1, 1).Obj(2).OBJIndex = PocionAmarillaNoCae
Recompensas(GUERRERO, 1, 1).Obj(2).Amount = 100
Recompensas(GUERRERO, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(GUERRERO, 1, 2).Obj(1).Amount = 1000
Recompensas(GUERRERO, 2, 1).SubeHP = 5

Recompensas(PIRATA, 1, 1).SubeHP = 20
Recompensas(PIRATA, 2, 2).SubeHP = 40
End Sub
Sub EstablecerRestas()

Resta(CIUDADANO) = 3
AumentoHit(CIUDADANO) = 3
Resta(TRABAJADOR) = 2.5
AumentoHit(TRABAJADOR) = 3
Resta(EXPERTO_MINERALES) = 2.5
AumentoHit(EXPERTO_MINERALES) = 3
Resta(MINERO) = 2.5
AumentoHit(MINERO) = 2
Resta(HERRERO) = 2.5
AumentoHit(HERRERO) = 2
Resta(EXPERTO_MADERA) = 2.5
AumentoHit(EXPERTO_MADERA) = 3
Resta(TALADOR) = 2.5
AumentoHit(TALADOR) = 2
Resta(CARPINTERO) = 2.5
AumentoHit(CARPINTERO) = 2
Resta(PESCADOR) = 2.5
AumentoHit(PESCADOR) = 1
Resta(SASTRE) = 2.5
AumentoHit(SASTRE) = 2
Resta(ALQUIMISTA) = 2.5
AumentoHit(ALQUIMISTA) = 2
Resta(LUCHADOR) = 3
AumentoHit(LUCHADOR) = 3
Resta(CON_MANA) = 3
AumentoHit(CON_MANA) = 3
Resta(HECHICERO) = 3
AumentoHit(HECHICERO) = 3
Resta(MAGO) = 3
AumentoHit(MAGO) = 1
Resta(NIGROMANTE) = 3
AumentoHit(NIGROMANTE) = 1
Resta(ORDEN_SAGRADA) = 1.5
AumentoHit(ORDEN_SAGRADA) = 3
Resta(PALADIN) = 0.5
AumentoHit(PALADIN) = 3
Resta(CLERIGO) = 1.5
AumentoHit(CLERIGO) = 2
Resta(NATURALISTA) = 2.5
AumentoHit(NATURALISTA) = 3
Resta(BARDO) = 1.5
AumentoHit(BARDO) = 2
Resta(Druida) = 3
AumentoHit(Druida) = 1.5
Resta(SIGILOSO) = 1.5
AumentoHit(SIGILOSO) = 3
Resta(ASESINO) = 1.5
AumentoHit(ASESINO) = 3
Resta(CAZADOR) = 0.5
AumentoHit(CAZADOR) = 3
Resta(SIN_MANA) = 2
AumentoHit(SIN_MANA) = 2
AumentoHit(ARQUERO) = 3
AumentoHit(GUERRERO) = 3
AumentoHit(CABALLERO) = 3
AumentoHit(BANDIDO) = 2
Resta(PIRATA) = 1.5
AumentoHit(PIRATA) = 2
Resta(LADRON) = 2.5
AumentoHit(LADRON) = 2

End Sub
Sub LoadMensajes()

Mensajes(Real, 1) = "||&HFF8080°¡No eres fiel al rey!°"
Mensajes(Real, 2) = "||&HFF8080°¡¡Maldito insolente!! ¡Los seguidores de Lord Thek no tienen lugar en nuestro ejército!°"
Mensajes(Real, 3) = "||&HFF8080°Tu Clan no responde a la Alianza del Fúrius, debes retirarte de él para poder enlistarte.°"
Mensajes(Real, 4) = "||&HFF8080°¡Ya perteneces a las tropas reales! ¡Ve a combatir criminales!°"
Mensajes(Real, 5) = "||&HFF8080°¡Para unirte a nuestras fuerzas debes matar al menos 40 criminales, solo has matado "
Mensajes(Real, 6) = "||&HFF8080°¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!°"
Mensajes(Real, 7) = "||&HFF8080°¡¡Bienvenido a al Ejército Imperial!! Si demuestras fidelidad al rey y destreza en las peleas, podrás aumentar de jerarquía.°"
Mensajes(Real, 8) = "%4"
Mensajes(Real, 9) = "5&"
Mensajes(Real, 10) = "8&"
Mensajes(Real, 11) = "N0"
Mensajes(Real, 12) = "L0"
Mensajes(Real, 13) = "J0"
Mensajes(Real, 14) = "K0"
Mensajes(Real, 15) = "M0"
Mensajes(Real, 16) = "||&HFF8080°¡No perteneces a las tropas reales!°"
Mensajes(Real, 17) = "||&HFF8080°Tu deber es combatir criminales, cada 50 criminales que derrotes te dare una recompensa.°"
Mensajes(Real, 18) = "||&HFF8080°¿Has decidido abandonarnos? Bien, ya nunca volveremos a aceptarte como ciudadano.°"
Mensajes(Real, 19) = "1W"
Mensajes(Real, 20) = "||Si ambos juraron fidelidad a la Alianza tienen que estar en clanes enemigos para poder atacarse." & FONTTYPE_FIGHT
Mensajes(Real, 21) = "/E"
Mensajes(Real, 22) = "||&HFF8080°¡Ya haz alcanzado la jerarquia más alta en las filas de la Alianza del Fúrius!°"
Mensajes(Real, 23) = "||&HFF8080°¡No puedes abandonar la Alianza del Fúrius! Perteneces a un clan ya, debes abandonarlo primero.°"

Mensajes(Caos, 1) = "||&H8080FF°¡No eres fiel a Lord Thek!°"
Mensajes(Caos, 2) = "||&H8080FF°¡¡Maldito insolente!! ¡Los seguidores del rey no tienen lugar en nuestro ejército!°"
Mensajes(Caos, 3) = "||&H8080FF°Tu Clan no responde al Ejército de Lord Thek, debes retirarte de él para poder enlistarte.°"
Mensajes(Caos, 4) = "||&H8080FF°¡Ya perteneces a las tropas del mal! ¡Ve a combatir ciudadanos!°"
Mensajes(Caos, 5) = "||&H8080FF°¡Para unirte a nuestras fuerzas debes matar al menos 40 ciudadanos, solo has matado "
Mensajes(Caos, 6) = "||&H8080FF°¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!°"
Mensajes(Caos, 7) = "||&H8080FF°¡Bienvenido al Ejército de Lord Thek! Si demuestras tu fidelidad y destreza en las peleas, podrás aumentar de jerarquía.°"
Mensajes(Caos, 8) = "%5"
Mensajes(Caos, 9) = "6&"
Mensajes(Caos, 10) = "9&"
Mensajes(Caos, 11) = "R0"
Mensajes(Caos, 12) = "P0"
Mensajes(Caos, 13) = "Ñ0"
Mensajes(Caos, 14) = "O0"
Mensajes(Caos, 15) = "Q0"
Mensajes(Caos, 16) = "||&H8080FF°¡No perteneces al Ejército de Lord Thek!°"
Mensajes(Caos, 17) = "||&H8080FF°Tu deber es combatir ciudadanos, cada 50 ciudadanos que derrotes te dare una recompensa.°"
Mensajes(Caos, 18) = "||&H8080FF°¡Traidor! ¡Jamás podrás volver con nosotros!°"
Mensajes(Caos, 19) = "2&"
Mensajes(Caos, 20) = "||Si ambos son seguidores de Lord Thek tienen que estar en clanes enemigos para poder atacarse." & FONTTYPE_FIGHT
Mensajes(Caos, 21) = "/D"
Mensajes(Caos, 22) = "||&H8080FF°¡Ya haz alcanzado la jerarquia más alta en las filas del Ejército de Lord Thek!°"
Mensajes(Caos, 23) = "||&H8080FF°¡No puedes abandonar el Ejército de Lord Thek! Perteneces a un clan ya, debes abandonarlo primero.°"

End Sub

Sub RevisarCarpetas()

If Not FileExist(App.Path & "\Logs", vbDirectory) Then Call MkDir$(App.Path & "\Logs")
If Not FileExist(App.Path & "\Logs\Consejeros", vbDirectory) Then Call MkDir$(App.Path & "\Logs\Consejeros")
If Not FileExist(App.Path & "\Logs\Data", vbDirectory) Then Call MkDir$(App.Path & "\Logs\Data")
If Not FileExist(App.Path & "\Foros", vbDirectory) Then Call MkDir$(App.Path & "\Foros")
If Not FileExist(App.Path & "\Guilds", vbDirectory) Then Call MkDir$(App.Path & "\Guilds")
If Not FileExist(App.Path & "\WorldBackUp", vbDirectory) Then Call MkDir$(App.Path & "\WorldBackUp")
If FileExist(App.Path & "\Logs\NPCs.log", vbNormal) Then Call Kill(App.Path & "\Logs\NPCs.log")

End Sub
Sub Listas()
Dim i As Integer

LevelSkill(1).LevelValue = 3
LevelSkill(2).LevelValue = 5
LevelSkill(3).LevelValue = 7
LevelSkill(4).LevelValue = 10
LevelSkill(5).LevelValue = 13
LevelSkill(6).LevelValue = 15
LevelSkill(7).LevelValue = 17
LevelSkill(8).LevelValue = 20
LevelSkill(9).LevelValue = 23
LevelSkill(10).LevelValue = 25
LevelSkill(11).LevelValue = 27
LevelSkill(12).LevelValue = 30
LevelSkill(13).LevelValue = 33
LevelSkill(14).LevelValue = 35
LevelSkill(15).LevelValue = 37
LevelSkill(16).LevelValue = 40
LevelSkill(17).LevelValue = 43
LevelSkill(18).LevelValue = 45
LevelSkill(19).LevelValue = 47
LevelSkill(20).LevelValue = 50
LevelSkill(21).LevelValue = 53
LevelSkill(22).LevelValue = 55
LevelSkill(23).LevelValue = 57
LevelSkill(24).LevelValue = 60
LevelSkill(25).LevelValue = 63
LevelSkill(26).LevelValue = 65
LevelSkill(27).LevelValue = 67
LevelSkill(28).LevelValue = 70
LevelSkill(29).LevelValue = 73
LevelSkill(30).LevelValue = 75
LevelSkill(31).LevelValue = 77
LevelSkill(32).LevelValue = 80
LevelSkill(33).LevelValue = 83
LevelSkill(34).LevelValue = 85
LevelSkill(35).LevelValue = 87
LevelSkill(36).LevelValue = 90
LevelSkill(37).LevelValue = 93
LevelSkill(38).LevelValue = 95
LevelSkill(39).LevelValue = 97
LevelSkill(40).LevelValue = 100
LevelSkill(41).LevelValue = 100
LevelSkill(42).LevelValue = 100
LevelSkill(43).LevelValue = 100
LevelSkill(44).LevelValue = 100
LevelSkill(45).LevelValue = 100
LevelSkill(46).LevelValue = 100
LevelSkill(47).LevelValue = 100
LevelSkill(48).LevelValue = 100
LevelSkill(49).LevelValue = 100
LevelSkill(50).LevelValue = 100

ELUs(1) = 300

For i = 2 To 10
    ELUs(i) = ELUs(i - 1) * 1.5
Next

For i = 11 To 24
    ELUs(i) = ELUs(i - 1) * 1.3
Next

For i = 25 To STAT_MAXELV - 1
    ELUs(i) = ELUs(i - 1) * 1.2
Next

'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Enano"
ListaRazas(3) = "Elfo"
ListaRazas(4) = "Elfo oscuro"
ListaRazas(5) = "Gnomo"

ReDim ListaBandos(0 To 2) As String
ListaBandos(0) = "Neutral"
ListaBandos(1) = "Alianza del Furius"
ListaBandos(2) = "Ejército de Lord Thek"

'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim ListaClases(1 To NUMCLASES) As String
ListaClases(1) = "Ciudadano"
ListaClases(2) = "Trabajador"
ListaClases(3) = "Experto en minerales"
ListaClases(4) = "Minero"
ListaClases(8) = "Herrero"
ListaClases(13) = "Experto en uso de madera"
ListaClases(14) = "Leñador"
ListaClases(18) = "Carpintero"
ListaClases(23) = "Pescador"
ListaClases(27) = "Sastre"
ListaClases(31) = "Alquimista"
ListaClases(35) = "Luchador"
ListaClases(36) = "Con uso de mana"
ListaClases(37) = "Hechicero"
ListaClases(38) = "Mago"
ListaClases(39) = "Nigromante"
ListaClases(40) = "Orden sagrada"
ListaClases(41) = "Paladin"
ListaClases(42) = "Clerigo"
ListaClases(43) = "Naturalista"
ListaClases(44) = "Bardo"
ListaClases(45) = "Druida"
ListaClases(46) = "Sigiloso"
ListaClases(47) = "Asesino"
ListaClases(48) = "Cazador"
ListaClases(49) = "Sin uso de mana"
ListaClases(50) = "Arquero"
ListaClases(51) = "Guerrero"
ListaClases(52) = "Caballero"
ListaClases(53) = "Bandido"
ListaClases(55) = "Pirata"
ListaClases(56) = "Ladron"

'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim SkillsNames(1 To NUMSKILLS) As String

SkillsNames(1) = "Magia"
SkillsNames(2) = "Robar"
SkillsNames(3) = "Tacticas de combate"
SkillsNames(4) = "Combate con armas"
SkillsNames(5) = "Meditar"
SkillsNames(6) = "Destreza con dagas"
SkillsNames(7) = "Ocultarse"
SkillsNames(8) = "Supervivencia"
SkillsNames(9) = "Talar árboles"
SkillsNames(10) = "Defensa con escudos"
SkillsNames(11) = "Pesca"
SkillsNames(12) = "Mineria"
SkillsNames(13) = "Carpinteria"
SkillsNames(14) = "Herreria"
SkillsNames(15) = "Liderazgo"
SkillsNames(16) = "Domar animales"
SkillsNames(17) = "Armas de proyectiles"
SkillsNames(18) = "Wresterling"
SkillsNames(19) = "Navegacion"
SkillsNames(20) = "Sastrería"
SkillsNames(21) = "Comercio"
SkillsNames(22) = "Resistencia Mágica"
SkillsNames(23) = "Botanica"
SkillsNames(24) = "Alquimia"


'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim UserSkills(1 To NUMSKILLS) As Integer

'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"

End Sub

Public Sub LoadGuildsNew()
Dim NumGuilds As Integer, GuildNum As Integer
Dim i As Integer, Num As Integer
Dim A As Long, S As Long
Dim NewGuild As cGuild


Dim Carpeta As String
Carpeta = App.Path & "\Guilds\"


If Not FileExist(App.Path & "\Guilds\GuildsInfo.inf", vbNormal) Then Exit Sub

A = INICarga(App.Path & "\Guilds\GuildsInfo.inf")
Call INIConf(A, 0, "", 0)

S = INIBuscarSeccion(A, "INIT")
NumGuilds = INIDarClaveInt(A, S, "NroGuilds")

For GuildNum = 1 To NumGuilds
    
   ' S = INIBuscarSeccion(A, "Guild" & GuildNum)

A = INICarga(Carpeta & "GUILD" & GuildNum & ".cla")
S = INIBuscarSeccion(A, "MAIN")

    If S >= 0 Then
        Set NewGuild = New cGuild
        With NewGuild
        .GuildName = INIDarClaveStr(A, S, "GuildName")
        .Founder = INIDarClaveStr(A, S, "Founder")
        .FundationDate = INIDarClaveStr(A, S, "Date")
        .Description = INIDarClaveStr(A, S, "Desc")
        
        .Codex = INIDarClaveStr(A, S, "Codex")
        
        .Leader = INIDarClaveStr(A, S, "Leader")
        .Gold = INIDarClaveInt(A, S, "Gold")
        .URL = INIDarClaveStr(A, S, "URL")
        .GuildExperience = INIDarClaveInt(A, S, "Exp")
        .DaysSinceLastElection = INIDarClaveInt(A, S, "DaysLast")
        .GuildNews = INIDarClaveStr(A, S, "GuildNews")
        .Bando = INIDarClaveInt(A, S, "Bando")
        
        'Num = INIDarClaveInt(A, S, "NumAliados")
       '
       ' For i = 1 To Num
       '     Call .AlliedGuilds.Add(INIDarClaveStr(A, S, "Aliado" & i))
       ' Next
        
       ' Num = INIDarClaveInt(A, S, "NumEnemigos")
        
       ' For i = 1 To Num
       '     Call .EnemyGuilds.Add(INIDarClaveStr(A, S, "Enemigo" & i))
       ' Next
        
        Num = INIDarClaveInt(A, S, "NumMiembros")
        
        For i = 1 To Num
            Call .Members.Add(INIDarClaveStr(A, S, "Miembro" & i))
        Next
        
        Num = INIDarClaveInt(A, S, "NumSolicitudes")
        
        Dim sol As cSolicitud
    
        For i = 1 To Num
            Set sol = New cSolicitud
            sol.UserName = ReadField(1, INIDarClaveStr(A, S, "Solicitud" & i), 172)
            sol.Desc = ReadField(2, INIDarClaveStr(A, S, "Solicitud" & i), 172)
            Call .Solicitudes.Add(sol)
        Next
        
        Num = INIDarClaveInt(A, S, "NumProposiciones")
        
        For i = 1 To Num
            Set sol = New cSolicitud
            sol.UserName = ReadField(1, INIDarClaveStr(A, S, "Proposicion" & i), 172)
            sol.Desc = ReadField(2, INIDarClaveStr(A, S, "Proposicion" & i), 172)
            Call .PeacePropositions.Add(sol)
        Next
        
        Call Guilds.Add(NewGuild)
        End With
    End If
Next

End Sub
Sub Main()
On Error Resume Next

Call Randomize(Timer)

ChDir App.Path
ChDrive App.Path

Call RevisarCarpetas
Call LoadMotd

Prision.Map = 66
Libertad.Map = 66

Prision.X = 53
Prision.Y = 47
Libertad.X = 53
Libertad.Y = 68

CaptureTheFlag.BanderaCrimiPos.Map = 194
CaptureTheFlag.BanderaCrimiPos.X = 86
CaptureTheFlag.BanderaCrimiPos.Y = 52

CaptureTheFlag.BanderaCiudaPos.Map = 196
CaptureTheFlag.BanderaCiudaPos.X = 15
CaptureTheFlag.BanderaCiudaPos.Y = 52
Call IniciarPiramide


'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim Resta(1 To NUMCLASES) As Single
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim Recompensas(1 To NUMCLASES, 1 To 3, 1 To 2) As Recompensa
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim AumentoHit(1 To NUMCLASES) As Byte
Call EstablecerRestas

LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")

'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim Npclist(1 To MAXNPCS) As Npc
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim CharList(1 To MAXCHARS) As Integer

IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"
MapPath = App.Path & "\Maps\"
MapDatFile = MapPath & "Info.dat"

Call Listas

frmCargando.Show

Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

'FIXIT: App.Revision property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
ENDL = Chr$(13) & Chr$(10)
ENDC = Chr$(1)
IniPath = App.Path & "\"
CharPath = App.Path & "\charfile\"

MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents
Call LoadSoportes
Call LoadBans
frmCargando.Label1(2).Caption = "Iniciando Arrays..."
Call LoadGuildsNew
Call CargarMods
Call CargarSpawnList
Call CargarForbidenWords
frmCargando.Label1(2).Caption = "Cargando Server.ini"
Call LoadSini
Call CargaNpcsDat
frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
Call LoadOBJData
Call LoadTops(Nivel)
Call LoadTops(Muertos)
Call LoadTops(RGanadas)

Call LoadMensajes
frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
Call CargarHechizos
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadEscudosHerreria
Call LoadCascosHerreria
Call LoadObjCarpintero
Call LoadObjSastre
Call LoadObjBotanica
Call LoadVentas
If MySql Then
Call CargarDB
End If
Call EstablecerRecompensas
Call LoadCasino
frmCargando.Label1(2).Caption = "Cargando Mapas"
Call LoadMapDataNew
If BootDelBackUp Then
    frmCargando.Label1(2).Caption = "Cargando BackUp"
    Call CargarBackUp
End If

Dim LoopC As Integer

NpcNoIniciado.Name = "NPC SIN INICIAR"
UserOffline.ConnID = -1
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
Next

If ClientsCommandsQueue = 1 Then
frmMain.CmdExec.Enabled = True
Else
frmMain.CmdExec.Enabled = False
End If

'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then
    Call IniciaWsApi
    SockListen = ListenForConnect(Puerto, hWndMsg, "")
'FIXIT: '#ElseIf' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#ElseIf UsarQueSocket = 0 Then
    
    frmCargando.Label1(2).Caption = "Configurando Sockets"
    
    frmMain.Socket2(0).AddressFamily = AF_INET
    frmMain.Socket2(0).protocol = IPPROTO_IP
    frmMain.Socket2(0).SocketType = SOCK_STREAM
    frmMain.Socket2(0).Binary = False
    frmMain.Socket2(0).Blocking = False
    frmMain.Socket2(0).BufferSize = 2048
    
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

Call NpcCanAttack(True)
Call NpcAITimer(True)
Call AutoTimer(True)

Unload frmCargando

Call LogMain("Server iniciado.")

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

tInicioServer = Timer
Call InicializaEstadisticas
ItemsUvU = True
End Sub
Public Sub ApagarSistema()
On Error GoTo Terminar
Dim UI As Integer

Call WorldSave
Call SaveGuildsNew
For UI = 1 To LastUser
    Call CloseSocket(UI)
Next
Call SaveSoportes
Call DescargaNpcsDat
Call NpcCanAttack(False)
Call NpcAITimer(False)
Call AutoTimer(False)
Call CerrarDB

Terminar:
End

End Sub
Function FileExist(file As String, FileType As VbFileAttribute) As Boolean

FileExist = Len(Dir$(file, FileType))

End Function
Public Function Tilde(Data As String) As String

Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(Data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")

End Function
Public Function ReadField(POS As Integer, Text As String, SepASCII As Integer) As String
Dim i As Integer, LastPos As Integer, FieldNum As Integer
If Len(Text) > 32000 Then Exit Function
For i = 1 To Len(Text)
    If Mid$(Text, i, 1) = Chr$(SepASCII) Then
        FieldNum = FieldNum + 1
        If FieldNum = POS Then
            ReadField = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(SepASCII), vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next

If FieldNum + 1 = POS Then ReadField = Mid$(Text, LastPos + 1)

End Function
Function MapaValido(Map As Integer) As Boolean

MapaValido = Map >= 1 And Map <= NumMaps

End Function
Sub MostrarNumUsers()

frmMain.CantUsuarios.Caption = NumNoGMs
Call SendData(ToAll, 0, 0, "NON" & NumNoGMs)

End Sub
Public Sub LogCriticEvent(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub LogBando(Bando As Byte, Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile
Select Case Bando
    Case Real
        Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
    Case Caos
        Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
End Select
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub LogMain(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\Logs\Main.log" For Append Shared As #nfile
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, Date & " " & Time, Desc
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub Logear(Archivo As String, Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\Logs\" & Archivo & ".log" For Append Shared As #nfile
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, Date & " " & Time, Desc
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub LogErrorUrgente(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\ErroresUrgentes.log" For Append Shared As #nfile
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub LogError(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogPENA(UserNombre As String, Texto As String, UserGM As Integer)
On Error GoTo Iafue
Dim NPenas As Integer
NPenas = val(GetVar(CharPath & "/" & UCase$(UserNombre) & ".chr", "PENAS", "CANTIDAD"))
Call WriteVar(CharPath & "/" & UCase$(UserNombre) & ".chr", "PENAS", "PENA" & NPenas + 1, UserList(UserGM).Name & "- " & Texto)
Call WriteVar(CharPath & "/" & UCase$(UserNombre) & ".chr", "PENAS", "CANTIDAD", NPenas + 1)
Exit Sub
Iafue:
Call LogError("ERROR EN LOGPENAAA NO SE Q ONDIS")
End Sub






Public Sub LogGM(Nombre As String, Texto As String, Consejero As Boolean)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile



If Consejero Then
    Open App.Path & "\logs\consejeros\" & Nombre & ".log" For Append Shared As #nfile
Else
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
End If
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, Date & " " & Time & " " & Texto
Close #nfile

Exit Sub

errhandler:

End Sub

'Public Sub SaveDayStats()
'On Error GoTo errhandler
'Dim nfile As Integer
'nfile = FreeFile
'Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile
'Print #nfile, "<stats>"
'Print #nfile, "<ao>"
'Print #nfile, "<dia>" & Date & "</dia>"
'Print #nfile, "<hora>" & Time & "</hora>"
'Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
'Print #nfile, "<max_user>" & DayStats.MaxUsuarios & "</max_user>"
'Print #nfile, "</ao>"
'Print #nfile, "</stats>"
'
'Close #nfile
'Exit Sub
'errhandler:

'End Sub

Public Sub LogVentaCasa(ByVal Texto As String)
'
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile

Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, "----------------------------------------------------------"
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, Date & " " & Time & " " & Texto
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub AutoBan(Texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\logs\AutoBan.log" For Append Shared As #nfile
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, "----------------------------------------------------------"
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, Date & " " & Time & " " & Texto
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub LogHackAttemp(Texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, "----------------------------------------------------------"
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, Date & " " & Time & " " & Texto
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub
Function ValidInputNP(cad As String) As Boolean
Dim Arg As String, i As Integer

For i = 1 To 33
    Arg = ReadField(i, cad, 44)
    If Len(Arg) = 0 Then Exit Function
Next

ValidInputNP = True

End Function
Sub Recargar()
Dim i As Integer

Call SendData(ToAll, 0, 0, "!!Recargando información, espere unos momentos.")

For i = 1 To LastUser
    Call CloseSocket(i)
Next

'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim Npclist(1 To MAXNPCS) As Npc
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim CharList(1 To MAXCHARS) As Integer

Recargando = True

Call CargarSpawnList
Call LoadSini
Call CargaNpcsDat
Call LoadOBJData
Call CargarHechizos
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadEscudosHerreria
Call LoadCascosHerreria
Call LoadObjCarpintero
Call LoadObjSastre
Call LoadMapDataNew
If BootDelBackUp Then Call CargarBackUp

For i = 1 To MaxUsers
    UserList(i).ConnID = -1
Next

Recargando = False

End Sub
Sub Restart()

On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

Dim LoopC As Integer
  
For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next
  
LastUser = 0
NumUsers = 0
NumNoGMs = 0

'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim Npclist(1 To MAXNPCS) As Npc
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData
Call LoadMapDataNew

Call CargarHechizos

'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 0 Then
    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup
    
    
    frmMain.Socket1.AddressFamily = AF_INET
    frmMain.Socket1.protocol = IPPROTO_IP
    frmMain.Socket1.SocketType = SOCK_STREAM
    frmMain.Socket1.Binary = False
    frmMain.Socket1.Blocking = False
    frmMain.Socket1.BufferSize = 1024
    
    frmMain.Socket2(0).AddressFamily = AF_INET
    frmMain.Socket2(0).protocol = IPPROTO_IP
    frmMain.Socket2(0).SocketType = SOCK_STREAM
    frmMain.Socket2(0).Blocking = False
    frmMain.Socket2(0).BufferSize = 2048
    
    
    frmMain.Socket1.LocalPort = val(Puerto)
    frmMain.Socket1.listen
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."


Call LogMain(" Servidor reiniciado.")



If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  
End Sub
Public Function Intemperie(userindex As Integer) As Boolean
    
If MapInfo(UserList(userindex).POS.Map).Zona <> "DUNGEON" Then
    Intemperie = MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y).trigger <> 1 And _
                MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y).trigger <> 2 And _
                MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y).trigger <> 4
End If
    
End Function
'Sub Desmontar(userindex As Integer)
'Dim Posss As WorldPos

'UserList(userindex).flags.Montado = 0
'Call Tilelibre(UserList(userindex).POS, Posss)
'Call TraerCaballo(userindex, UserList(userindex).flags.CaballoMontado + 1, Posss.X, Posss.Y, Posss.Map)
'UserList(userindex).flags.CaballoMontado = -1
'UserList(userindex).Char.Head = UserList(userindex).OrigChar.Head
'If UserList(userindex).Invent.ArmourEqpObjIndex Then
'    UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Ropaje
'Else
'    Call DarCuerpoDesnudo(userindex)
'End If
'If UserList(userindex).Invent.EscudoEqpObjIndex Then _
'    UserList(userindex).Char.ShieldAnim = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).ShieldAnim
'If UserList(userindex).Invent.WeaponEqpObjIndex Then _
'    UserList(userindex).Char.WeaponAnim = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).WeaponAnim
'If UserList(userindex).Invent.CascoEqpObjIndex Then _
'    UserList(userindex).Char.CascoAnim = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).CascoAnim
'Call ChangeUserChar(ToMap, 0, UserList(userindex).POS.Map, userindex, RopaEquitacion(userindex), UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
'Call SendData(ToIndex, userindex, 0, "MONTA0")
'End Sub
'Function RopaEquitacion(userindex As Integer) As Integer

'If RazaBaja(userindex) Then
'    RopaEquitacion = ROPA_DE_EQUITACION_ENANO
'Else
'    RopaEquitacion = ROPA_DE_EQUITACION_NORMAL
'End If

'End Function
'Public Sub TraerCaballo(userindex As Integer, ByVal Num As Integer, Optional X As Integer, Optional Y As Integer, Optional Map As Integer)
'Dim NPCNN As Integer
'Dim Poss As WorldPos
'If Map Then
'    Poss.Map = Map
'    Poss.X = X
'    Poss.Y = Y
'Else
'    Poss = Ubicar(UserList(userindex).POS)
'End If

'NPCNN = SpawnNpc(108, Poss, False, False)

'UserList(userindex).Caballos.NpcNum(Num - 1) = NPCNN
'UserList(userindex).Caballos.POS(Num - 1) = Npclist(NPCNN).POS'

'End Sub
Public Function Ubicar(POS As WorldPos) As WorldPos
On Error GoTo errhandler

Dim NuevaPos As WorldPos
NuevaPos.X = 0
NuevaPos.Y = 0
Call Tilelibre(POS, NuevaPos)
If NuevaPos.X <> 0 And NuevaPos.Y Then
    Ubicar = NuevaPos
End If

Exit Function

errhandler:
End Function
'Sub QuitarCaballos(userindex As Integer)
'Dim i As Integer '

'For i = 0 To UserList(userindex).Caballos.Num - 1
'    QuitarNPC (UserList(userindex).Caballos.NpcNum(i))
'Next

'End Sub
Public Sub CargaNpcsDat()
Dim npcfile As String

npcfile = DatPath & "NPCs.dat"
ANpc = INICarga(npcfile)
Call INIConf(ANpc, 0, "", 0)

npcfile = DatPath & "NPCs-HOSTILES.dat"
Anpc_host = INICarga(npcfile)
Call INIConf(Anpc_host, 0, "", 0)

End Sub
Public Sub DescargaNpcsDat()

If ANpc Then Call INIDescarga(ANpc)
If Anpc_host Then Call INIDescarga(Anpc_host)

End Sub
Sub GuardarUsuarios()
Dim i As Integer

Call SendData(ToAll, 0, 0, "2R")
If MySql Then

For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then Call SaveUserSQL(i)
Next i

Else
    For i = 1 To LastUser
        If UserList(i).flags.EnDM = False Then
        If UserList(i).flags.UserLogged Then Call SaveUser(i, CharPath & UCase$(UserList(i).Name) & ".chr")
        End If
    Next i
    
End If

Call SendData(ToAll, 0, 0, "3R")

End Sub

''Call SaveFurius
Sub GuardarUsuariosFurius()
Dim i As Integer

Call SendData(ToAll, 0, 0, "2R")
    For i = 1 To LastUser
        If UserList(i).flags.EnDM = False Then
        If UserList(i).flags.UserLogged Then Call SaveFurius(i, CharPath & UCase$(UserList(i).Name) & ".chr")
        End If
    Next i
Call SendData(ToAll, 0, 0, "3R")

End Sub

Sub InicializaEstadisticas()

Call EstadisticasWeb.Inicializa(frmMain.hwnd)
Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call EstadisticasWeb.Informar(UPTIME_SERVER, (TiempoTranscurrido(tInicioServer)) / 1000)
Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub
