Attribute VB_Name = "ModFacciones"
Option Explicit
Public Sub Recompensado(UserIndex As Integer)
Dim Fuerzas As Byte
Dim MiObj As Obj

Fuerzas = UserList(UserIndex).Faccion.Bando


If UserList(UserIndex).Faccion.Jerarquia = 0 Then
    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 11))
    Exit Sub
End If

If UserList(UserIndex).Faccion.Jerarquia = 1 Then
    If UserList(UserIndex).Faccion.Matados(Enemigo(Fuerzas)) < 90 Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 12) & 500)
        Exit Sub
    End If
    

    UserList(UserIndex).Faccion.Jerarquia = 2
    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 15) & Titulo(UserIndex))
ElseIf UserList(UserIndex).Faccion.Jerarquia = 2 Then
    If UserList(UserIndex).Faccion.Matados(Enemigo(Fuerzas)) < 140 Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 12) & 140)
        Exit Sub
    End If
    
  
    
    UserList(UserIndex).Faccion.Jerarquia = 3
    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 15) & Titulo(UserIndex))
ElseIf UserList(UserIndex).Faccion.Jerarquia = 3 Then
    If UserList(UserIndex).Faccion.Matados(Enemigo(Fuerzas)) < 190 Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 12) & 190)
        Exit Sub
    End If
    
 
    
  
   
    UserList(UserIndex).Faccion.Jerarquia = 4
    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 15) & Titulo(UserIndex))
End If


If UserList(UserIndex).Faccion.Jerarquia < 4 Then
    MiObj.Amount = 1
    MiObj.OBJIndex = Armaduras(Fuerzas, UserList(UserIndex).Faccion.Jerarquia, TipoClase(UserIndex), TipoRaza(UserIndex))
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)
Else
    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 22) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
End If

End Sub
Public Sub Expulsar(UserIndex As Integer)

Call SendData(ToIndex, UserIndex, 0, Mensajes(UserList(UserIndex).Faccion.Bando, 8))
UserList(UserIndex).Faccion.Bando = Neutral
Call UpdateUserChar(UserIndex)

End Sub
Public Sub Enlistar(UserIndex As Integer, ByVal Fuerzas As Byte)
Dim MiObj As Obj

If UserList(UserIndex).Faccion.Bando = Neutral Then
    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 1) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.Bando = Enemigo(Fuerzas) Then
    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 2) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 Then
    If oGuild.Bando <> Fuerzas Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 3) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If
End If

If UserList(UserIndex).Faccion.Jerarquia Then
    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 4) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.Matados(Enemigo(Fuerzas)) < 40 Then
    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 5) & UserList(UserIndex).Faccion.Matados(Enemigo(Fuerzas)) & "!�" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 6) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 7) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

UserList(UserIndex).Faccion.Jerarquia = 1

MiObj.Amount = 1
MiObj.OBJIndex = Armaduras(Fuerzas, UserList(UserIndex).Faccion.Jerarquia, TipoClase(UserIndex), TipoRaza(UserIndex))
If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj)

Call LogBando(Fuerzas, UserList(UserIndex).Name)

End Sub
Public Function Titulo(UserIndex As Integer) As String

Select Case UserList(UserIndex).Faccion.Bando
    Case Real
        Select Case UserList(UserIndex).Faccion.Jerarquia
            Case 0
                Titulo = "Fiel al Rey"
            Case 1
                Titulo = "Soldado Real"
            Case 2
                Titulo = "General Real"
            Case 3
                Titulo = "Elite Real"
            Case 4
                Titulo = "Pr�tector de Newbies"
        End Select
    Case Caos
        Select Case UserList(UserIndex).Faccion.Jerarquia
            Case 0
                Titulo = "Fiel a Lord Thek"
            Case 1
                Titulo = "Ac�lito"
            Case 2
                Titulo = "Jefe de Tropas"
            Case 3
                Titulo = "Elite del Mal"
            Case 4
                Titulo = "H�roe del Mal"
        End Select
End Select


End Function
