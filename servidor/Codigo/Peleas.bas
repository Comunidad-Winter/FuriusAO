Attribute VB_Name = "Peleas"
Public APUESTATOTAL As Integer
Public APUESTAMAXIMA As Integer
Public Iniciada As Boolean



Public Sub NuevaApuesta(userindex As Integer, apuesta As Integer)
If APUESTATOTAL + apuesta > APUESTAMAXIMA Then

Exit Sub
End If

End Sub

Public Sub OfrecerPelea(userindex As Integer, Valorx As Integer)


If Valorx < 1000 Then
Call SendData(ToIndex, userindex, 0, "||No puedes apostar menos de 1000 monedas de oro" & FONTTYPE_VENENO)
Exit Sub
End If

If UserList(userindex).Stats.GLD < Valorx Then
Call SendData(ToIndex, userindex, 0, "||No tienes oro suficiente" & FONTTYPE_VENENO)
Exit Sub
End If

APUESTAMAXIMA = Valorx * 2
Call SendData(ToAll, 0, 0, "||" & UserList(userindex).Name & " está ofreciendo pelea por el monto de " & Valorx & FONTTYPE_VENENO)

End Sub


Public Sub IniciarPelea(userindex As Integer)
If UserList(userindex).Stats.GLD < APUESTAMAXIMA / 2 Then
Call SendData(ToIndex, userindex, 0, "||No tienes el oro suficiente" & FONTTYPE_VENENO)
Exit Sub
End If


UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - APUESTAMAXIMA / 2
Iniciada = True
'aceptador.oro = apuestamaxima / 2


End Sub
