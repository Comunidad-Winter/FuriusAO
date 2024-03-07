Attribute VB_Name = "Extra"
'fúriusao 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@furiusao.com.ar
'www.furiusao.com.ar

Option Explicit
Public Function EsNewbie(userindex As Integer) As Boolean

EsNewbie = (UserList(userindex).Stats.ELV <= LimiteNewbie)

End Function
Public Sub DoTileEvents(userindex As Integer)
On Error GoTo errhandler
Dim Map As Integer, x As Integer, Y As Integer
Dim nPos As WorldPos, mPos As WorldPos

Map = UserList(userindex).POS.Map
x = UserList(userindex).POS.x
Y = UserList(userindex).POS.Y

If MapData(Map, x, Y).trigger = 9 And UserList(userindex).flags.Muerto = 1 Then Call RevivirUsuarioNPC(userindex): Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & ",36,1")
If MapData(Map, x, Y).trigger = 10 And UserList(userindex).flags.Muerto = 0 Then Call CompruebaBloques
'CompruebaBloques
mPos = MapData(Map, x, Y).TileExit
If Not MapaValido(mPos.Map) Or Not InMapBounds(mPos.x, mPos.Y) Then Exit Sub




If MapInfo(mPos.Map).Restringir And Not EsNewbie(userindex) Then
    Call SendData(ToIndex, userindex, 0, "1J")
ElseIf UserList(userindex).Stats.ELV < MapInfo(mPos.Map).Nivel And Not (UserList(userindex).Clase = PIRATA And UserList(userindex).Recompensas(1) = 2) Then
    Call SendData(ToIndex, userindex, 0, "%/" & MapInfo(mPos.Map).Nivel)
Else
    If LegalPos(mPos.Map, mPos.x, mPos.Y, PuedeAtravesarAgua(userindex), UserList(userindex).flags.Privilegios) Then
        If mPos.x <> 0 And mPos.Y <> 0 Then Call WarpUserChar(userindex, mPos.Map, mPos.x, mPos.Y, ObjData(MapData(Map, x, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT)
    Else
        Call ClosestStablePos(mPos, nPos)
        If nPos.x <> 0 And nPos.Y Then Call WarpUserChar(userindex, nPos.Map, nPos.x, nPos.Y, ObjData(MapData(Map, x, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT)
    End If
    Exit Sub
End If

Call ClosestStablePos(UserList(userindex).POS, nPos)
If nPos.x <> 0 And nPos.Y Then Call WarpUserChar(userindex, nPos.Map, nPos.x, nPos.Y, ObjData(MapData(Map, x, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT)

Exit Sub

errhandler:
    Call LogError("Error en DoTileEvents-" & nPos.Map & "-" & nPos.x & "-" & nPos.Y)

End Sub
Function InMapBounds(x As Integer, Y As Integer) As Boolean

InMapBounds = (x >= MinXBorder And x <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder)

End Function
Sub ClosestStablePos(POS As WorldPos, ByRef nPos As WorldPos)
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = POS.Map

Do While Not LegalPos(POS.Map, nPos.x, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = POS.Y - LoopC To POS.Y + LoopC
        For tX = POS.x - LoopC To POS.x + LoopC
            
            If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                nPos.x = tX
                nPos.Y = tY

                tX = POS.x + LoopC
                tY = POS.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.Y = 0
End If

End Sub
Sub ClosestLegalPos(POS As WorldPos, nPos As WorldPos, Optional AguaValida As Boolean)
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = POS.Map

Do While Not LegalPos(POS.Map, nPos.x, nPos.Y, AguaValida)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = POS.Y - LoopC To POS.Y + LoopC
        For tX = POS.x - LoopC To POS.x + LoopC
            
            If LegalPos(nPos.Map, tX, tY, AguaValida) Then
                nPos.x = tX
                nPos.Y = tY
                
                
                tX = POS.x + LoopC
                tY = POS.Y + LoopC
  
            End If
        
        Next
    Next
    
    LoopC = LoopC + 1
    
Loop

If Notfound Then
    nPos.x = 0
    nPos.Y = 0
End If

End Sub
Function ClaseIndex(ByVal Clase As String) As Integer
Dim i As Integer

For i = 1 To UBound(ListaClases)
    If UCase$(ListaClases(i)) = UCase$(Clase) Then
        ClaseIndex = i
        Exit Function
    End If
Next

End Function
Function NameIndex(ByVal Name As String) As Integer
Dim userindex As Integer, i As Integer

Name = Replace$(Name, "+", " ")

If Len(Name) = 0 Then
    NameIndex = 0
    Exit Function
End If
  
userindex = 1

If Right$(Name, 1) = "*" Then
    Name = Left$(Name, Len(Name) - 1)
    For i = 1 To LastUser
        If UCase$(UserList(i).Name) = UCase$(Name) Then
            NameIndex = i
            Exit Function
        End If
    Next
Else
    For i = 1 To LastUser
        If UCase$(Left$(UserList(i).Name, Len(Name))) = UCase$(Name) Then
            NameIndex = i
            Exit Function
        End If
    Next
End If

End Function
Function CheckForSameIP(userindex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged Then
        If UserList(LoopC).ip = UserIP And userindex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next

End Function
Function CheckForSameName(userindex As Integer, ByVal Name As String) As Boolean
Dim LoopC As Integer

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged Then
        If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next

End Function



Function CheckForSamePC(ByVal PCL As String) As Boolean
Dim LoopC As Integer

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged Then
        If UCase$(UserList(LoopC).flags.PCLabel) = UCase$(PCL) Then
            CheckForSamePC = True
            Exit Function
        End If
    End If
Next

End Function



Sub HeadtoPos(Head As Byte, POS As WorldPos)
Dim x As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

x = POS.x
Y = POS.Y

If Head = NORTH Then
    nX = x
    nY = Y - 1
End If

If Head = SOUTH Then
    nX = x
    nY = Y + 1
End If

If Head = EAST Then
    nX = x + 1
    nY = Y
End If

If Head = WEST Then
    nX = x - 1
    nY = Y
End If

POS.x = nX
POS.Y = nY

End Sub
Function LegalPos(Map As Integer, x As Integer, Y As Integer, Optional PuedeAgua As Boolean, Optional UserExIndex As Integer) As Boolean

If Not MapaValido(Map) Or Not InMapBounds(x, Y) Then Exit Function

LegalPos = (MapData(Map, x, Y).Blocked = 0) And _
           (MapData(Map, x, Y).userindex = 0) And _
           (MapData(Map, x, Y).NpcIndex = 0) And _
           (MapData(Map, x, Y).Agua = Buleano(PuedeAgua) Or val(UserExIndex) > 0)
           

End Function
Function LegalPosNPC(Map As Integer, x As Integer, Y As Integer, AguaValida As Boolean) As Boolean

If Not InMapBounds(x, Y) Then Exit Function

LegalPosNPC = (MapData(Map, x, Y).Blocked <> 1) And _
     (MapData(Map, x, Y).userindex = 0) And _
     (MapData(Map, x, Y).NpcIndex = 0) And _
     (MapData(Map, x, Y).trigger <> POSINVALIDB) And _
     (MapData(Map, x, Y).trigger <> POSINVALIDA) _
     And Buleano(AguaValida) = MapData(Map, x, Y).Agua
     
End Function
Public Sub SendNPC(userindex As Integer, NpcIndex As Integer)
Dim Info As String
Dim CRI As Byte

Select Case UserList(userindex).Stats.UserSkills(Supervivencia)
    Case Is <= 20
        If Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP Then
            CRI = 5
        Else: CRI = 1
        End If
    Case Is < 40
        Select Case 100 * Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP
            Case 100
                CRI = 5
            Case Is >= 50
                CRI = 2
            Case Else
                CRI = 3
        End Select
    Case Is < 60
        Select Case 100 * Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP
            Case 100
                CRI = 5
            Case Is > 66
                CRI = 2
            Case Is > 33
                CRI = 3
            Case Else
                CRI = 4
        End Select
    Case Is < 100
        CRI = 5 + Fix(10 * (1 - (Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP)))
    Case Else
        Info = "||" & Npclist(NpcIndex).Name & " [" & Npclist(NpcIndex).Stats.MinHP & "/" & Npclist(NpcIndex).Stats.MaxHP & "]"
        If Npclist(NpcIndex).flags.Paralizado Then Info = Info & " - PARALIZADO"
        Call SendData(ToIndex, userindex, 0, Info & FONTTYPE_INFO)
        Exit Sub
End Select

Info = "9Q" & Npclist(NpcIndex).Name & "," & CRI
Call SendData(ToIndex, userindex, 0, Info)
                
End Sub
Public Sub Expresar(NpcIndex As Integer, userindex As Integer)

If Npclist(NpcIndex).NroExpresiones Then
'FIXIT: Declare 'randomi' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
    Dim randomi
    randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "3Q" & vbWhite & "°" & Npclist(NpcIndex).Expresiones(randomi) & "°" & Npclist(NpcIndex).Char.CharIndex)
End If
                    
End Sub
Sub LookatTile(userindex As Integer, Map As Integer, x As Integer, Y As Integer)

Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim NPMUERTO As String
Dim Info As String


If InMapBounds(x, Y) Then
    UserList(userindex).flags.TargetMap = Map
    UserList(userindex).flags.TargetX = x
    UserList(userindex).flags.TargetY = Y
    
    If MapData(Map, x, Y).OBJInfo.OBJIndex Then
        
        If MapData(Map, x, Y).OBJInfo.Amount = 1 Then
            Call SendData(ToIndex, userindex, 0, "4Q" & ObjData(MapData(Map, x, Y).OBJInfo.OBJIndex).Name)
        Else
            Call SendData(ToIndex, userindex, 0, "5Q" & ObjData(MapData(Map, x, Y).OBJInfo.OBJIndex).Name & "," & MapData(Map, x, Y).OBJInfo.Amount)
        End If
        UserList(userindex).flags.TargetObj = MapData(Map, x, Y).OBJInfo.OBJIndex
        UserList(userindex).flags.TargetObjMap = Map
        UserList(userindex).flags.TargetObjX = x
        UserList(userindex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(Map, x + 1, Y).OBJInfo.OBJIndex Then
        
        If ObjData(MapData(Map, x + 1, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
            Call SendData(ToIndex, userindex, 0, "6Q" & ObjData(MapData(Map, x + 1, Y).OBJInfo.OBJIndex).Name)
            UserList(userindex).flags.TargetObj = MapData(Map, x + 1, Y).OBJInfo.OBJIndex
            UserList(userindex).flags.TargetObjMap = Map
            UserList(userindex).flags.TargetObjX = x + 1
            UserList(userindex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, x + 1, Y + 1).OBJInfo.OBJIndex Then
        If ObjData(MapData(Map, x + 1, Y + 1).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
            
            Call SendData(ToIndex, userindex, 0, "6Q" & ObjData(MapData(Map, x + 1, Y + 1).OBJInfo.OBJIndex).Name)
            UserList(userindex).flags.TargetObj = MapData(Map, x + 1, Y + 1).OBJInfo.OBJIndex
            UserList(userindex).flags.TargetObjMap = Map
            UserList(userindex).flags.TargetObjX = x + 1
            UserList(userindex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, x, Y + 1).OBJInfo.OBJIndex Then
        If ObjData(MapData(Map, x, Y + 1).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
            
            Call SendData(ToIndex, userindex, 0, "6Q" & ObjData(MapData(Map, x, Y + 1).OBJInfo.OBJIndex).Name)
            UserList(userindex).flags.TargetObj = MapData(Map, x, Y).OBJInfo.OBJIndex
            UserList(userindex).flags.TargetObjMap = Map
            UserList(userindex).flags.TargetObjX = x
            UserList(userindex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    
    If Y + 1 <= YMaxMapSize Then
        If MapData(Map, x, Y + 1).userindex Then
            TempCharIndex = MapData(Map, x, Y + 1).userindex
            FoundChar = 1
        End If
        If MapData(Map, x, Y + 1).NpcIndex Then
            TempCharIndex = MapData(Map, x, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    
    If FoundChar = 0 Then
        If MapData(Map, x, Y).userindex Then
            TempCharIndex = MapData(Map, x, Y).userindex
            FoundChar = 1
        End If
        If MapData(Map, x, Y).NpcIndex Then
            TempCharIndex = MapData(Map, x, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    
    If FoundChar = 1 Then
            
        If UserList(TempCharIndex).flags.AdminInvisible Then Exit Sub
        
        If UserList(TempCharIndex).Faccion.Bando Then
            If UserList(TempCharIndex).Faccion.BandoOriginal <> UserList(TempCharIndex).Faccion.Bando Then
                Stat = Stat & " <" & ListaBandos(UserList(TempCharIndex).Faccion.Bando) & "> <Mercenario>"
            ElseIf UserList(TempCharIndex).Faccion.Jerarquia Then
                Stat = Stat & " <" & ListaBandos(UserList(TempCharIndex).Faccion.Bando) & "> <" & Titulo(TempCharIndex) & ">"
            Else
                Stat = Stat & " <" & Titulo(TempCharIndex) & ">"
            End If
        End If
        
           ' If UserList(TempCharIndex).flags.Casado <> "" Then
          '  Dim ReRaRo$
          '  ReRaRo$ = IIf(UCase$(UserList(TempCharIndex).Genero) = "HOMBRE", "Casado", "Casada")
          '      Stat = Stat & " <" & ReRaRo$ & " con " & UserList(TempCharIndex).flags.Casado & ">"
         '   End If
        
        If Len(UserList(TempCharIndex).GuildInfo.GuildName) > 0 Then
            Stat = Stat & " <" & UserList(TempCharIndex).GuildInfo.GuildName & ">"
        End If
        
        If Len(UserList(TempCharIndex).Desc) > 0 Then
            Stat = UserList(TempCharIndex).Name & Stat & " - " & UserList(TempCharIndex).Desc
        Else
            Stat = UserList(TempCharIndex).Name & Stat
        End If
         
        If UserList(TempCharIndex).flags.Silenciado > 0 Then
        Stat = Stat & "[SILENCIADO]"
        End If
        
        
        If UserList(TempCharIndex).flags.EnDM = True Then
        Stat = Stat & " [DEATHMATCH] [Kills: "
        Stat = Stat & UserList(TempCharIndex).flags.DmKills
        Stat = Stat & "] [Muertes: " & UserList(TempCharIndex).flags.DmMuertes & "]"
        End If
       
       
       
       
            Select Case UserList(TempCharIndex).flags.Privilegios
                Case 1
                Stat = Stat & " <Soporte>" '& Stat
                Case 2
                Stat = Stat & " <Game Master>"
                Case 3
                Stat = Stat & " <Coordinación General>" ' & Stat
                Case 4
                Stat = Stat & " <Administración>" '& Stat
            End Select
        
        If UserList(TempCharIndex).flags.Privilegios Then
            Stat = "9J" & Stat
        End If
        If UserList(TempCharIndex).flags.Privilegios = 0 Then
        If UserList(TempCharIndex).flags.Muerto Then
                Stat = "2K" & UserList(TempCharIndex).Name
            ElseIf UserList(TempCharIndex).Faccion.Bando = Real Then
                Stat = "3K" & Stat
            ElseIf UserList(TempCharIndex).Faccion.Bando = Caos Then
                Stat = "4K" & Stat
            ElseIf EsNewbie(TempCharIndex) Then
                Stat = "H0" & Stat
            Else
                Stat = "1&" & Stat
            End If
            If UserList(TempCharIndex).flags.ConsejoCiuda Or UserList(TempCharIndex).flags.AyudanteCiuda Then
            Stat = Stat & " <Consejo de Banderbill>"
            ElseIf UserList(TempCharIndex).flags.ConsejoCaoz Or UserList(TempCharIndex).flags.AyudanteCaoz Then
            Stat = Stat & " <Concilio de Arghal>"
            End If
            
        End If
        
        If UserList(userindex).flags.Privilegios > 1 Then Stat = Stat & " [FPS:" & UserList(TempCharIndex).flags.Fps & "]"
        If UserList(userindex).flags.Privilegios > 1 Then Stat = Stat & " [Nivel:" & UserList(TempCharIndex).Stats.ELV & "]"
        Call SendData(ToIndex, userindex, 0, Stat)
           
         
        
        FoundSomething = 1
        UserList(userindex).flags.TargetUser = TempCharIndex
        UserList(userindex).flags.TargetNpc = 0
        UserList(userindex).flags.TargetNpcTipo = 0
       
       
    ElseIf FoundChar = 2 Then
            
            Dim wPos As WorldPos
            wPos.Map = Map
            wPos.x = x
            wPos.Y = Y
            If Distancia(Npclist(TempCharIndex).POS, wPos) > 1 Then
                MapData(Map, x, Y).NpcIndex = 0
                Exit Sub
            End If
                
            If Npclist(TempCharIndex).flags.TiendaUser Then
                If userindex = Npclist(TempCharIndex).flags.TiendaUser Then
                    If UserList(userindex).Tienda.Gold Then
                        Call SendData(ToIndex, userindex, 0, "/O" & UserList(userindex).Tienda.Gold & "," & Npclist(TempCharIndex).Char.CharIndex)
                    Else
                        Call SendData(ToIndex, userindex, 0, "/P" & Npclist(TempCharIndex).Char.CharIndex)
                    End If
                Else
                    Call SendData(ToIndex, userindex, 0, "/Q" & UserList(Npclist(TempCharIndex).flags.TiendaUser).Name & "," & Npclist(TempCharIndex).Char.CharIndex)
                End If
            ElseIf Len(Npclist(TempCharIndex).Desc) > 1 Then
                Call SendData(ToIndex, userindex, 0, "3Q" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).Char.CharIndex)
            ElseIf Npclist(TempCharIndex).MaestroUser Then
                Call SendData(ToIndex, userindex, 0, "7Q" & Npclist(TempCharIndex).Name & "," & UserList(Npclist(TempCharIndex).MaestroUser).Name)
            ElseIf Npclist(TempCharIndex).AutoCurar = 1 Then
                Call SendData(ToIndex, userindex, 0, "8Q" & Npclist(TempCharIndex).Name)
            Else
                Call SendNPC(userindex, TempCharIndex)
            End If
            FoundSomething = 1
            UserList(userindex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(userindex).flags.TargetNpc = TempCharIndex
            UserList(userindex).flags.TargetUser = 0
            UserList(userindex).flags.TargetObj = 0
    End If
    
    If FoundChar = 0 Then
        UserList(userindex).flags.TargetNpc = 0
        UserList(userindex).flags.TargetNpcTipo = 0
        UserList(userindex).flags.TargetUser = 0
    End If
    
    If FoundSomething = 0 Then
        UserList(userindex).flags.TargetNpc = 0
        UserList(userindex).flags.TargetNpcTipo = 0
        UserList(userindex).flags.TargetUser = 0
        UserList(userindex).flags.TargetObj = 0
        UserList(userindex).flags.TargetObjMap = 0
        UserList(userindex).flags.TargetObjX = 0
        UserList(userindex).flags.TargetObjY = 0
    End If

Else
    If FoundSomething = 0 Then
        UserList(userindex).flags.TargetNpc = 0
        UserList(userindex).flags.TargetNpcTipo = 0
        UserList(userindex).flags.TargetUser = 0
        UserList(userindex).flags.TargetObj = 0
        UserList(userindex).flags.TargetObjMap = 0
        UserList(userindex).flags.TargetObjX = 0
        UserList(userindex).flags.TargetObjY = 0
    End If
End If

End Sub
Function FindDirection(POS As WorldPos, Target As WorldPos) As Byte
Dim x As Integer, Y As Integer

x = POS.x - Target.x
Y = POS.Y - Target.Y

If Sgn(x) = -1 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

If Sgn(x) = 1 And Sgn(Y) = 1 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(x) = 1 And Sgn(Y) = -1 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(x) = -1 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

If Sgn(x) = 0 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

If Sgn(x) = 0 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

If Sgn(x) = 1 And Sgn(Y) = 0 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(x) = -1 And Sgn(Y) = 0 Then
    FindDirection = EAST
    Exit Function
End If

If Sgn(x) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function
Public Function ItemEsDeMapa(ByVal Map As Integer, x As Integer, Y As Integer) As Boolean

If MapData(Map, x, Y).OBJInfo.OBJIndex = FOGATA Then
ItemEsDeMapa = False
Exit Function
End If


ItemEsDeMapa = ObjData(MapData(Map, x, Y).OBJInfo.OBJIndex).Agarrable Or MapData(Map, x, Y).Blocked

End Function

