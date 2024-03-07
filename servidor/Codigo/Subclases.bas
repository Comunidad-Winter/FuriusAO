Attribute VB_Name = "modSubclases"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Sub EnviarSubclase(userindex As Integer)

If PuedeSubirClase(userindex) Then Call SendData(ToIndex, userindex, 0, "RECOM" & UserList(userindex).Clase)

End Sub
Public Sub EnviarRecom(userindex As Integer)
Dim cad As String
cad = cad & UserList(userindex).Clase & ","
cad = cad & UserList(userindex).Recompensas(1) & ","
cad = cad & UserList(userindex).Recompensas(2) & ","
cad = cad & UserList(userindex).Recompensas(3)

Call SendData(ToIndex, userindex, 0, "REC0" & cad)
End Sub
Sub RecibirRecompensa(userindex As Integer, Eleccion As Byte)
Dim Recompensa As Byte
Dim i As Integer

Recompensa = PuedeRecompensa(userindex)

If Recompensa = 0 Then Exit Sub

UserList(userindex).Recompensas(Recompensa) = Eleccion

If Recompensas(UserList(userindex).Clase, Recompensa, Eleccion).SubeHP Then
    Call AddtoVar(UserList(userindex).Stats.MaxHP, Recompensas(UserList(userindex).Clase, Recompensa, Eleccion).SubeHP, STAT_MAXHP)
    Call SendUserMAXHP(userindex)
End If

If Recompensas(UserList(userindex).Clase, Recompensa, Eleccion).SubeMP Then
    Call AddtoVar(UserList(userindex).Stats.MaxMAN, Recompensas(UserList(userindex).Clase, Recompensa, Eleccion).SubeMP, 2000 + 200 * Buleano(UserList(userindex).Clase = MAGO) * 200 + 300 * Buleano(UserList(userindex).Clase = MAGO And UserList(userindex).Recompensas(2) = 2))
    Call SendUserMAXMANA(userindex)
End If

For i = 1 To 2
    If Recompensas(UserList(userindex).Clase, Recompensa, Eleccion).Obj(i).OBJIndex Then
        If Not MeterItemEnInventario(userindex, Recompensas(UserList(userindex).Clase, Recompensa, Eleccion).Obj(i)) Then Call TirarItemAlPiso(UserList(userindex).POS, Recompensas(UserList(userindex).Clase, Recompensa, Eleccion).Obj(i))
    End If
Next

If PuedeRecompensa(userindex) = 0 Then Call SendData(ToIndex, userindex, 0, "SURE0")

End Sub
Sub RecibirSubclase(Clase As Byte, userindex As Integer)

If Not PuedeSubirClase(userindex) Then Exit Sub

Select Case UserList(userindex).Clase
    Case CIUDADANO
        If Clase = 1 Then
            UserList(userindex).Clase = TRABAJADOR
        Else: UserList(userindex).Clase = LUCHADOR
        End If

    Case TRABAJADOR
        Select Case Clase
            Case 1
                UserList(userindex).Clase = EXPERTO_MINERALES
            Case 2
                UserList(userindex).Clase = EXPERTO_MADERA
            Case 3
                UserList(userindex).Clase = PESCADOR
            Case 4
                UserList(userindex).Clase = SASTRE
        End Select
        
    Case EXPERTO_MINERALES
        If Clase = 1 Then
            UserList(userindex).Clase = MINERO
        Else: UserList(userindex).Clase = HERRERO
        End If
        
    Case EXPERTO_MADERA
        If Clase = 1 Then
            UserList(userindex).Clase = TALADOR
        Else: UserList(userindex).Clase = CARPINTERO
        End If
        
    Case LUCHADOR
        If Clase = 1 Then
            UserList(userindex).Clase = CON_MANA
            Call Aprenderhechizo(userindex, 2)
            UserList(userindex).Stats.MaxMAN = 100
            Call SendUserMAXMANA(userindex)
            If Not PuedeSubirClase(userindex) Then Call SendData(ToIndex, userindex, 0, "SUCL0")
            Exit Sub
        Else: UserList(userindex).Clase = SIN_MANA
        End If
        
    Case CON_MANA
        Select Case Clase
            Case 1
                UserList(userindex).Clase = HECHICERO
            Case 2
                UserList(userindex).Clase = ORDEN_SAGRADA
            Case 3
                UserList(userindex).Clase = NATURALISTA
            Case 4
                UserList(userindex).Clase = SIGILOSO
        End Select
        
    Case HECHICERO
        If Clase = 1 Then
            UserList(userindex).Clase = MAGO
        Else: UserList(userindex).Clase = NIGROMANTE
        End If

    Case ORDEN_SAGRADA
        If Clase = 1 Then
            UserList(userindex).Clase = PALADIN
        Else
            UserList(userindex).Clase = CLERIGO
        End If
    
    Case NATURALISTA
        If Clase = 1 Then
            UserList(userindex).Clase = BARDO
        Else: UserList(userindex).Clase = DRUIDA
        End If
        
    Case SIGILOSO
        If Clase = 1 Then
            UserList(userindex).Clase = ASESINO
        Else: UserList(userindex).Clase = CAZADOR
        End If
        
    Case SIN_MANA
        If Clase = 1 Then
            UserList(userindex).Clase = BANDIDO
        Else: UserList(userindex).Clase = CABALLERO
        End If
        
    Case BANDIDO
        If Clase = 1 Then
            UserList(userindex).Clase = PIRATA
        Else: UserList(userindex).Clase = LADRON
        End If
        
    Case CABALLERO
        If Clase = 1 Then
            UserList(userindex).Clase = GUERRERO
        Else: UserList(userindex).Clase = ARQUERO
        End If
End Select

Call CalcularValores(userindex)
Call SendData(ToIndex, userindex, 0, "/0" & ListaClases(UserList(userindex).Clase))

If Not PuedeSubirClase(userindex) Then Call SendData(ToIndex, userindex, 0, "SUCL0")

End Sub
