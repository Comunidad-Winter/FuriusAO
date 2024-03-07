Attribute VB_Name = "Rey"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Public QuienConquista(5) As String


Public Sub IrCastillo(userindex As Integer, Castillo As Integer)

Dim i As Double

For i = 1 To LastNPC
    If Npclist(i).ReyC = Castillo Then
        If QuienConquista(Npclist(i).ReyC) = UserList(userindex).GuildInfo.GuildName Then
            Call WarpUserChar(userindex, Npclist(i).POS.Map, Npclist(i).POS.X + 1, Npclist(i).POS.Y, False)
            Exit Sub
        End If
    End If
DoEvents
Next i

End Sub
