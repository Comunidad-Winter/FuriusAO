Attribute VB_Name = "Piramide"
Public Bloque(4) As Bloke
Type Bloke
    Map As Integer
    x As Integer
    Y As Integer
    'Activado As Boolean
End Type
Public PiramideActivada As Boolean
Public Sub IniciarPiramide()
PiramideActivada = True

'MAPAS'
'200 = ESTA LA MOMIA
'201 = 1 BLOQUE
'202 = 1 BLOQUE
'199 = 2 BLOQUES
'
'
'
'
'
'

'INICIAMOS DESDE ACA?? YES.
Bloque(1).Map = 202
Bloque(1).x = 40
Bloque(1).Y = 26

Bloque(2).Map = 201
Bloque(2).x = 54
Bloque(2).Y = 10

Bloque(3).Map = 199
Bloque(3).x = 48
Bloque(3).Y = 58

Bloque(4).Map = 199
Bloque(4).x = 55
Bloque(4).Y = 58

Dim MomiaBloque As WorldPos
MomiaBloque.Map = 200
MomiaBloque.x = 50
MomiaBloque.Y = 50
Call SpawnNpc(654, MomiaBloque, True, True)





End Sub
Public Sub CompruebaBloques()
On Error GoTo errrr
Dim A As Integer

For A = 1 To 4
    If MapData(Bloque(A).Map, Bloque(A).x, Bloque(A).Y).userindex = 0 Then
        Exit Sub
    Else
        If UserList(MapData(Bloque(A).Map, Bloque(A).x, Bloque(A).Y).userindex).flags.Muerto = 1 Then Exit Sub
    End If
    Call SendData(ToIndex, MapData(Bloque(A).Map, Bloque(A).x, Bloque(A).Y).userindex, 0, "||El bloque numero " & A & " está activado" & FONTTYPE_BLANCO)
Next A

Dim BC As Integer
Dim UserPCh(4) As Integer
UserPCh(1) = MapData(Bloque(1).Map, Bloque(1).x, Bloque(1).Y).userindex
UserPCh(2) = MapData(Bloque(2).Map, Bloque(2).x, Bloque(2).Y).userindex
UserPCh(3) = MapData(Bloque(3).Map, Bloque(3).x, Bloque(3).Y).userindex
UserPCh(4) = MapData(Bloque(4).Map, Bloque(4).x, Bloque(4).Y).userindex

Dim HPirata As Byte
Dim HClan As Byte
Dim HFaccion As Byte
Dim FOriginal As Byte

Dim Xx As Byte
For Xx = 1 To 4
If UserList(UserPCh(Xx)).Clase = PIRATA Then
If UserList(UserPCh(Xx)).Stats.ELV > 39 Then HPirata = 1: Exit For
End If
Next Xx


Xx = 0
For Xx = 1 To 4
If Len(UserList(UserPCh(Xx)).GuildInfo.GuildName) > 0 Then HClan = 1: Exit For
Next Xx

FOriginal = UserList(UserPCh(1)).Faccion.Bando
For Xx = 2 To 4
If (UserList(UserPCh(Xx)).Faccion.Bando) <> FOriginal Then HFaccion = 1: Exit For
Next Xx


If HPirata = 0 Then
    For BC = 1 To 4
    Call SendData(ToIndex, UserPCh(BC), 0, "||Debe haber al menos un pirata de nivel 40 o superior!" & FONTTYPE_INFO)
    Next BC
    Exit Sub
End If

If HClan = 1 Then
    For BC = 1 To 4
    Call SendData(ToIndex, UserPCh(BC), 0, "||Ninguno de los luchadores debe poseer clan." & FONTTYPE_INFO)
    Next BC
    Exit Sub
End If

If HFaccion = 1 Then
    For BC = 1 To 4
    Call SendData(ToIndex, UserPCh(BC), 0, "||Deben ser todos del mismo bando." & FONTTYPE_INFO)
    Next BC
    Exit Sub
End If



For BC = 1 To 4
'If MapData(Bloque(BC).Map, Bloque(BC).x, Bloque(BC).Y).userindex > 0 Then
Call WarpUserChar(UserPCh(BC), 199, 49 + BC, 69, True)
'End If
Next BC

Exit Sub
errrr:
Call LogError("Error en CompruebaBloques sub " & Err.Description)
Debug.Print Err.Description
End Sub
