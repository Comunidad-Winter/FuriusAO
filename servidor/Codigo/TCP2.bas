Attribute VB_Name = "TCP2"
Public Sub ActivarTrampa(userindex As Integer)
Dim i As Integer, TU As Integer

For i = 1 To MapInfo(UserList(userindex).POS.Map).NumUsers
    TU = MapInfo(UserList(userindex).POS.Map).userindex(i)
    If UserList(TU).flags.Paralizado = 0 And Abs(UserList(userindex).POS.x - UserList(TU).POS.x) <= 3 And Abs(UserList(userindex).POS.Y - UserList(TU).POS.Y) <= 3 And TU <> userindex And PuedeAtacar(userindex, TU) Then
       UserList(TU).flags.QuienParalizo = userindex
       UserList(TU).flags.Paralizado = 1
       UserList(TU).Counters.Paralisis = Timer - 15 * Buleano(UserList(TU).Clase = GUERRERO And UserList(TU).Recompensas(3) = 2)
       Call SendData(ToIndex, TU, 0, "PU" & UserList(TU).POS.x & "," & UserList(TU).POS.Y)
       Call SendData(ToIndex, TU, 0, ("P9"))
       Call SendData(ToPCArea, TU, UserList(TU).POS.Map, "CFX" & UserList(TU).Char.CharIndex & ",12,1")
    End If
Next

Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW112")

End Sub
Public Sub DesactivarMercenarios()
Dim userindex As Integer

For userindex = 1 To LastUser
    If UserList(userindex).Faccion.Bando <> Neutral And UserList(userindex).Faccion.Bando <> UserList(userindex).Faccion.BandoOriginal Then
        Call SendData(ToIndex, userindex, 0, "||La quest ha terminado, has dejado de ser un mercenario." & FONTTYPE_furius)
        UserList(userindex).Faccion.Bando = Neutral
        Call UpdateUserChar(userindex)
    End If
Next

End Sub
Public Function YaVigila(Espiado As Integer, Espiador As Integer) As Boolean
Dim i As Integer

For i = 1 To 10
    If UserList(Espiado).flags.Espiado(i) = Espiador Then
        UserList(Espiado).flags.Espiado(i) = 0
        YaVigila = True
        Exit Function
    End If
Next

End Function

Public Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
Next

AsciiValidos = True

End Function

Public Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next

Numeric = True

End Function
Public Function NombrePermitido(ByVal Nombre As String) As Boolean
Dim i As Integer

'FIXIT: Reexmplazar la función 'Trim' con la función 'Trim$'.                               FixIT90210ae-R9757-R1B8ZE
If Trim(Nombre) = "" Then NombrePermitido = False: Exit Function


For i = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(i)) Then
        NombrePermitido = False
        Exit Function
    End If
Next

NombrePermitido = True

End Function

Public Function ValidateAtrib(userindex As Integer) As Boolean
Dim LoopC As Integer

For LoopC = 1 To NUMATRIBUTOS
    If UserList(userindex).Stats.UserAtributosBackUP(LoopC) > 23 Or UserList(userindex).Stats.UserAtributosBackUP(LoopC) < 1 Then Exit Function
Next

ValidateAtrib = True

End Function

Public Function ValidateAtrib2(userindex As Integer) As Boolean
Dim LoopC As Integer

For LoopC = 1 To NUMATRIBUTOS
    If UserList(userindex).Stats.UserAtributosBackUP(LoopC) > 18 Or UserList(userindex).Stats.UserAtributosBackUP(LoopC) < 1 Then
    ValidateAtrib2 = False
    Exit Function
    End If
Next

ValidateAtrib2 = True

End Function
Public Function ValidateSkills(userindex As Integer) As Boolean
Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(userindex).Stats.UserSkills(LoopC) < 0 Then Exit Function
    If UserList(userindex).Stats.UserSkills(LoopC) > 100 Then UserList(userindex).Stats.UserSkills(LoopC) = 100
Next

ValidateSkills = True

End Function


