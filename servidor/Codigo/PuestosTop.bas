Attribute VB_Name = "PuestosTop"

Option Explicit
Public Function TotalMatados(userindex As Integer) As Integer

TotalMatados = UserList(userindex).Faccion.Matados(0) + UserList(userindex).Faccion.Matados(1) + UserList(userindex).Faccion.Matados(2)

End Function
Public Sub RevisarTops(userindex As Integer)

If UserList(userindex).flags.Privilegios > 0 Then
    If IndexTop(Nivel, userindex) <> UBound(Tops, 2) Then Call SacarTop(Nivel, userindex)
    If IndexTop(Muertos, userindex) <> UBound(Tops, 2) Then Call SacarTop(Muertos, userindex)
    If IndexTop(RGanadas, userindex) <> UBound(Tops, 2) Then Call SacarTop(RGanadas, userindex)
Else
    If UserList(userindex).Stats.ELV > Tops(Nivel, UBound(Tops, 2)).Nivel Then Call AgregarTop(Nivel, userindex)
    If TotalMatados(userindex) > Tops(Muertos, UBound(Tops, 2)).Muertos Then Call AgregarTop(Muertos, userindex)
    If UserList(userindex).flags.MatadasenR > Tops(RGanadas, UBound(Tops, 2)).RGanadas Then Call AgregarTop(RGanadas, userindex)
End If

End Sub
Public Function IndexTop(Top As Byte, userindex As Integer) As Integer
Dim i As Integer

For i = 1 To UBound(Tops, 2)
    If UCase$(Tops(Top, i).Nombre) = UCase$(UserList(userindex).Name) Then
        IndexTop = i
        Exit Function
    End If
Next

IndexTop = UBound(Tops, 2)

End Function
Public Sub AgregarTop(Top As Byte, userindex As Integer)
Dim i As Integer

i = IndexTop(Top, userindex)

For i = i - 1 To 1 Step -1
    If (Top = Nivel And UserList(userindex).Stats.ELV <= Tops(Nivel, i).Nivel) Or _
        (Top = Muertos And TotalMatados(userindex) <= Tops(Muertos, i).Muertos) Or _
        (Top = RGanadas And UserList(userindex).flags.MatadasenR <= Tops(RGanadas, i).RGanadas) Then
        i = i + 1
        Exit For
    End If
    Tops(Top, i + 1) = Tops(Top, i)
    Call SaveTop(Top, i + 1)
Next

i = Maximo(1, i)

Tops(Top, i).Nombre = UserList(userindex).Name
Tops(Top, i).Bando = ListaBandos(UserList(userindex).Faccion.Bando)
Tops(Top, i).Nivel = UserList(userindex).Stats.ELV
Tops(Top, i).Muertos = TotalMatados(userindex)
Tops(Top, i).RGanadas = UserList(userindex).flags.MatadasenR
Tops(Top, i).RPerdidas = UserList(userindex).flags.PerdidasenR
Call SaveTop(Top, i)

End Sub
Public Sub SacarTop(Top As Byte, userindex As Integer)
Dim i As Integer

i = IndexTop(Top, userindex)

For i = i To UBound(Tops, 2) - 1
    Tops(Top, i) = Tops(Top, i + 1)
    Call SaveTop(Top, i)
Next

Tops(Top, UBound(Tops, 2)).Nombre = ""
Tops(Top, UBound(Tops, 2)).Bando = ""
Tops(Top, UBound(Tops, 2)).Nivel = 0
Tops(Top, UBound(Tops, 2)).Muertos = 0
Tops(Top, UBound(Tops, 2)).RGanadas = 0
Tops(Top, UBound(Tops, 2)).RPerdidas = 0

Call SaveTop(Top, UBound(Tops, 2))

End Sub
Public Sub SaveTop(Top As Byte, Puesto As Integer)
Dim file As String
Dim i As Integer

If Len(Tops(Top, Puesto).Nombre) = 0 Then Exit Sub

Select Case Top
Case 1
file = App.Path & "\LOGS\TopNivel.log"
Case 2
file = App.Path & "\LOGS\TopMuertos.log"
Case 3
file = App.Path & "\LOGS\TopRetos.log"
End Select

Call WriteVar(file, "Top" & Puesto, "Name", Tops(Top, Puesto).Nombre)
Call WriteVar(file, "Top" & Puesto, "Nivel", val(Tops(Top, Puesto).Nivel))
Call WriteVar(file, "Top" & Puesto, "Muertos", val(Tops(Top, Puesto).Muertos))
Call WriteVar(file, "Top" & Puesto, "Bando", Tops(Top, Puesto).Bando)
Call WriteVar(file, "Top" & Puesto, "Ganadas", val(Tops(Top, Puesto).RGanadas))
Call WriteVar(file, "Top" & Puesto, "Perdidas", val(Tops(Top, Puesto).RPerdidas))

End Sub
Public Sub LoadTops(Top As Byte)
Dim file As String, i As Integer

Select Case Top
Case 1
file = App.Path & "\LOGS\TopNivel.log"
Case 2
file = App.Path & "\LOGS\TopMuertos.log"
Case 3
file = App.Path & "\LOGS\TopRetos.log"
End Select

If Not FileExist(file, vbNormal) Then Exit Sub

For i = 1 To UBound(Tops, 2)
    Tops(Top, i).Nombre = GetVar(file, "Top" & i, "Name")
    Tops(Top, i).Nivel = val(GetVar(file, "Top" & i, "Nivel"))
    Tops(Top, i).Muertos = val(GetVar(file, "Top" & i, "Muertos"))
    Tops(Top, i).Bando = GetVar(file, "Top" & i, "Bando")
    Tops(Top, i).RGanadas = GetVar(file, "Top" & i, "Ganadas")
    Tops(Top, i).RPerdidas = GetVar(file, "Top" & i, "Perdidas")
    
Next

End Sub

