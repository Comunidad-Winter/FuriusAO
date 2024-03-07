Attribute VB_Name = "Matematicas"

Option Explicit
'FIXIT: Declare 'Total' and 'Porc' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function Porcentaje(Total As Variant, Porc As Variant) As Long

Porcentaje = Total * (Porc / 100)

End Function
'FIXIT: Declare 'Var' and 'Take' and 'MIN' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Sub RestVar(Var As Variant, Take As Variant, MIN As Variant)

Var = Maximo(Var - Take, MIN)

End Sub
'FIXIT: Declare 'Var' and 'Addon' and 'MAX' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Sub AddtoVar(Var As Variant, Addon As Variant, MAX As Variant)

Var = Minimo(Var + Addon, MAX)

End Sub
'FIXIT: Declare 'Distancia' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Function Distancia(wp1 As WorldPos, wp2 As WorldPos)

Distancia = (Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100))

End Function
Function TipoClase(userindex As Integer) As Byte

Select Case UserList(userindex).Clase
    Case PALADIN, ASESINO, CAZADOR
        TipoClase = 2
    Case CLERIGO, BARDO, LADRON
        TipoClase = 3
    Case MAGO, NIGROMANTE, DRUIDA
        TipoClase = 4
    Case Else
        TipoClase = 1
End Select

End Function
Public Function TipoRaza(userindex As Integer) As Byte

If UserList(userindex).Raza = ENANO Or UserList(userindex).Raza = GNOMO Then
    TipoRaza = 2
Else: TipoRaza = 1
End If

End Function
Public Function RazaBaja(userindex As Integer) As Boolean

RazaBaja = (UserList(userindex).Raza = ENANO Or UserList(userindex).Raza = GNOMO)

End Function
Function Distance(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer) As Double

Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

End Function
