Attribute VB_Name = "Retos2vs2"
Public Type Parej
User1 As Integer
User2 As Integer
End Type


Public Pareja1 As Parej
Public Pareja2 As Parej

Sub SeFue(userindex As Integer)
On Error GoTo erri
If Reto2vs2EnCursO = True Then
Call DevolverParticipantes
Call SendData(ToAll, 0, 0, "||Ring 2> El reto ha sido cancelado por la desconexión de " & UserList(userindex).Name & FONTTYPE_BLANCO)
'.....'
Else
    If userindex = Pareja1.User1 Or userindex = Pareja1.User2 Then
    Pareja1.User1 = 0
    Pareja1.User2 = 0
    Else
    Pareja2.User1 = 0
    Pareja2.User2 = 0
    End If
End If
Exit Sub
erri:
Call LogError("Error en SeFue user:" & UserList(userindex).Name)

End Sub

Sub DevolverParticipantes()
Call WarpUserChar(Pareja1.User1, 160, 60, 35, True)
Call WarpUserChar(Pareja1.User2, 160, 61, 35, True)
Call WarpUserChar(Pareja2.User1, 160, 60, 36, True)
Call WarpUserChar(Pareja2.User2, 160, 61, 36, True)

UserList(Pareja1.User1).flags.EnReto = 0
UserList(Pareja1.User2).flags.EnReto = 0
UserList(Pareja2.User1).flags.EnReto = 0
UserList(Pareja2.User2).flags.EnReto = 0

UserList(Pareja1.User1).flags.Pareja = 0
UserList(Pareja1.User2).flags.Pareja = 0
UserList(Pareja2.User1).flags.Pareja = 0
UserList(Pareja2.User2).flags.Pareja = 0

UserList(Pareja1.User1).flags.Parejado = 0
UserList(Pareja1.User2).flags.Parejado = 0
UserList(Pareja2.User1).flags.Parejado = 0
UserList(Pareja2.User2).flags.Parejado = 0

Reto2vs2EnCursO = False
Call LimpiarParejas
End Sub


Sub LimpiarParejas()
Pareja1.User1 = 0
Pareja1.User2 = 0
Pareja2.User1 = 0
Pareja2.User2 = 0
End Sub

Public Sub LlevarParejas()
Call SendData(ToMap, 0, 160, "||Ring 2> " & UserList(Pareja1.User1).Name & " - " & UserList(Pareja1.User2).Name & " se enfrentan a " & UserList(Pareja2.User1).Name & " - " & UserList(Pareja2.User2).Name & FONTTYPE_BLANCO)
Call WarpUserChar(Pareja1.User1, 170, 12, 81, True)
Call WarpUserChar(Pareja1.User2, 170, 12, 80, True)
Call WarpUserChar(Pareja2.User1, 170, 25, 91, True)
Call WarpUserChar(Pareja2.User2, 170, 28, 90, True)
End Sub

Public Sub Pagar(ParejaNum As Integer)
DoEvents
If ParejaNum = 1 Then
UserList(Pareja1.User1).Stats.GLD = UserList(Pareja1.User1).Stats.GLD + 200000
UserList(Pareja1.User2).Stats.GLD = UserList(Pareja1.User2).Stats.GLD + 200000
Call SendUserStatsBox(Pareja1.User1)
Call SendUserStatsBox(Pareja1.User2)
ElseIf ParejaNum = 2 Then
UserList(Pareja2.User1).Stats.GLD = UserList(Pareja2.User1).Stats.GLD + 200000
UserList(Pareja2.User2).Stats.GLD = UserList(Pareja2.User2).Stats.GLD + 200000
Call SendUserStatsBox(Pareja2.User1)
Call SendUserStatsBox(Pareja2.User2)
End If
DoEvents
End Sub
