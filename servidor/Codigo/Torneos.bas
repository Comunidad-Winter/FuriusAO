Attribute VB_Name = "Torneos"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Public Type Participante
Nombre As String
'lugar As Integer
Indice As Integer
End Type

Public Type TorneoB
MAXPARTICIPANTES As Integer
ParticipantesTotales As Integer
Participantes(30) As Participante
Iniciado As Boolean
Cerrado As Boolean
Ronda As Integer
PrimerJugador As Integer
UltimoJugador As Integer
ClaseUnica As String
NivelMinimo As Integer
Precio As Long
End Type
Public Torneo As TorneoB

Public Blc As Integer


Public Sub IniciarTorneo()
'Dim Mapap, xP, Yp As Integer
'    Mapap = 191
'    xP = 64
'    Yp = 57
'    MapData(Mapap, xP, Yp).OBJInfo.OBJIndex = 0
 '   MapData(Mapap, xP, Yp).OBJInfo.Amount = 0
 '   Call SendData(ToMap, 0, 191, "BO" & xP & "," & Yp)
 '   MapData(Mapap, xP, Yp).TileExit.Map = 0
 ''   MapData(Mapap, xP, Yp).TileExit.X = 0
 '   MapData(Mapap, xP, Yp).TileExit.Y = 0
Call SendData(ToAll, 0, 0, "||Torneo > Un torneo automatico 1 Vs. 1 dará comienzo, /ENTRAR, Ganador queda en campo" & FONTTYPE_BLANCO)
Torneo.Iniciado = True
Torneo.Cerrado = False
Torneo.ParticipantesTotales = 0
Torneo.Ronda = 0
For Blc = 1 To Torneo.MAXPARTICIPANTES
Torneo.Participantes(Blc).Nombre = ""
'Torneo.Participantes(Blc).lugar = 0
Torneo.Participantes(Blc).Indice = 0
DoEvents
Next Blc
End Sub

Public Sub TerminarTorneo()
'Call SendData(ToAll, 0, 0, "||Un torneo automatico 1 vs 1 dará comienzo /ENTRAR" & FONTTYPE_BLANCO)
Torneo.Iniciado = False
Torneo.Cerrado = True
Torneo.ParticipantesTotales = 0
Torneo.Ronda = 0
Torneo.ClaseUnica = ""
'Dim mapad, Xd, Yd As Integer
   
  '  Dim ET As Obj
  '  ET.Amount = 1
  '  ET.OBJIndex = Teleport
    
 '   Call MakeObj(ToMap, 0, 191, ET, 191, 64, 57)
 '   MapData(191, 64, 57).TileExit.Map = 1
 '   MapData(191, 64, 57).TileExit.X = 50
 '   MapData(191, 64, 57).TileExit.Y = 50
    
    
    
Dim Blcx As Integer
For Blcx = 1 To Torneo.MAXPARTICIPANTES
Torneo.Participantes(Blc).Nombre = ""
'Torneo.Participantes(Blc).lugar = 0
Torneo.Participantes(Blcx).Indice = 0
If UserList(Torneo.Participantes(Blcx).Indice).flags.EnTorneo = True Then UserList(Torneo.Participantes(Blcx).Indice).flags.EnTorneo = False
DoEvents
Next Blcx
Call SendData(ToAdmins, 0, 0, "||Admins > Torneo terminado" & FONTTYPE_VENENO)
End Sub


Public Sub CambiarParticipanteS(Cantidad As Integer)
Torneo.MAXPARTICIPANTES = Cantidad
'ReDim Torneo.Participantes(cantidad)
End Sub



Public Sub InscribirUsuario(userindex As Integer)
If Torneo.ParticipantesTotales = Torneo.MAXPARTICIPANTES Then
Call SendData(ToIndex, userindex, 0, "||El cupo esta lleno" & FONTTYPE_CELESTE)
Exit Sub
End If
If UserList(userindex).flags.Muerto = 1 Then Exit Sub
If UserList(userindex).Counters.Pena <> 0 Then Exit Sub

If Torneo.Cerrado Then Exit Sub
If Not Torneo.Iniciado Then Exit Sub


If Torneo.ClaseUnica <> "" And Torneo.ClaseUnica <> "TODAS" Then
   If UCase$(ListaClases(UserList(userindex).Clase)) <> UCase$(Torneo.ClaseUnica) Then
    Call SendData(ToIndex, userindex, 0, "||Tu clase no puede participar de este torneo" & FONTTYPE_BLANCO)
    Exit Sub
    End If
End If


If Torneo.NivelMinimo > UserList(userindex).Stats.ELV Then
  Call SendData(ToIndex, userindex, 0, "||Tu nivel no te permite participar de este torneo" & FONTTYPE_BLANCO)
  Exit Sub
End If


Call SendData(ToIndex, userindex, 0, "||Estas inscripto" & FONTTYPE_VENENO)

Call WarpUserChar(userindex, 191, 50, 50, True)
Torneo.ParticipantesTotales = Torneo.ParticipantesTotales + 1
Torneo.Participantes(Torneo.ParticipantesTotales).Nombre = UserList(userindex).Name
Torneo.Participantes(Torneo.ParticipantesTotales).Indice = userindex




UserList(userindex).flags.EnTorneo = True
If Torneo.ParticipantesTotales = Torneo.MAXPARTICIPANTES Then
Torneo.Cerrado = True
Torneo.Iniciado = True
Torneo.PrimerJugador = 1
Torneo.UltimoJugador = 2
'Call WarpUserChar(Torneo.Participantes(1).Indice, 86, 50, 50, False)
Call Peleas
End If

End Sub

Public Sub PerdioRonda(userindex As Integer, Gano As Integer)

For Blc = 1 To Torneo.MAXPARTICIPANTES

If Torneo.Participantes(Blc).Indice = Gano Then
Torneo.PrimerJugador = Blc
End If


DoEvents
Next Blc
UserList(userindex).flags.EnTorneo = False
Call WarpUserChar(userindex, 1, 50, 50, True)

Call SendData(ToAll, 0, 0, "||Torneo > " & UserList(Gano).Name & " ha derrotado a " & UserList(userindex).Name & FONTTYPE_BLANCO)


If Torneo.PrimerJugador >= Torneo.MAXPARTICIPANTES Or userindex = Torneo.Participantes(Torneo.MAXPARTICIPANTES).Indice Then
Call SendData(ToAll, 0, 0, "||Torneo > " & UserList(Gano).Name & " ha ganado el torneo" & FONTTYPE_BLANCO)
Call WarpUserChar(Gano, 1, 50, 50, True)
UserList(Gano).flags.EnTorneo = False
TerminarTorneo
Exit Sub
End If



Torneo.UltimoJugador = Torneo.UltimoJugador + 1
Call Peleas
End Sub

Public Sub Peleas()
Torneo.Ronda = Torneo.Ronda + 1


If UserList(Torneo.Participantes(Torneo.UltimoJugador).Indice).flags.Muerto = 1 Or _
UserList(Torneo.Participantes(Torneo.UltimoJugador).Indice).Name <> Torneo.Participantes(Torneo.UltimoJugador).Nombre Or _
UserList(Torneo.Participantes(Torneo.UltimoJugador).Indice).flags.EnTorneo = False Then
Call PerdioRonda(Torneo.Participantes(Torneo.UltimoJugador).Indice, Torneo.Participantes(Torneo.PrimerJugador).Indice)
Exit Sub
End If

Call WarpUserChar(Torneo.Participantes(Torneo.PrimerJugador).Indice, 86, 38, 19)
Call WarpUserChar(Torneo.Participantes(Torneo.UltimoJugador).Indice, 86, 51, 28)
Call SendData(ToAll, 0, 0, "||Torneo > La ronda " & Torneo.Ronda & " ha dado comienzo! En esta se enfrentan " & Torneo.Participantes(Torneo.PrimerJugador).Nombre & " vs " & Torneo.Participantes(Torneo.UltimoJugador).Nombre & FONTTYPE_BLANCO)
End Sub
