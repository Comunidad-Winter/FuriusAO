Attribute VB_Name = "modSOPORTE"
Public conn As ADODB.Connection
Dim Bd As Boolean

Public Sub BDDResetGMsos()
    On Error GoTo mal
    If Bd = False Then Exit Sub
    conn.Execute "DELETE FROM sos"
    Exit Sub
mal:
End Sub

Public Sub BDDAddGMsos(user As String, razon As String)
    On Error GoTo mal
    If Bd = False Then Exit Sub
    conn.Execute "INSERT INTO sos VALUES('" & user & "','" & razon & "')"
    Exit Sub
mal:
End Sub

Public Sub BDDDelGMsos(user As String)
    On Error GoTo mal
    If Bd = False Then Exit Sub
    conn.Execute "DELETE FROM sos WHERE user='" & user & "'"
    Exit Sub
mal:
End Sub


Public Sub BDDConnect()
   On Error GoTo mal
'QUITAR ESTO
'GoTo mal
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "DRIVER={MySQL ODBC 3.51 Driver};" _
      & "SERVER=db.localhost.net.ar;" _
      & "DATABASE=furiusao;" _
      & "UID=furiusao;PWD=lentejuela; OPTION=3"
    conn.Open
    conn.Execute "DROP TABLE IF EXISTS sos"
    conn.Execute "CREATE TABLE sos(fecha text, mapa text, personaje text, email text, servidor text, gm text, asunto text, mensaje text, respondido text, censura text, old text, respondidopor text, respondidoel text, respuesta text)"
    Bd = True
    Exit Sub
    
mal:
End Sub
