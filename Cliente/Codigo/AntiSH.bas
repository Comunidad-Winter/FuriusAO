Attribute VB_Name = "AntiSH"
Option Explicit
'NIST SERVERS
'128.138.140.44    utcnist.colorado.edu           University of CO, Boulder
'129.6.15.28       time-a.nist.gov                NIST, Gaithersburg, MD
'129.6.15.29       time-b.nist.gov                NIST, Gaithersburg, MD
'131.107.1.10      time-nw.nist.gov               Microsoft, Redmond, WA
'132.163.4.101     time-a.timefreq.bldrdoc.gov    NIST, Boulder, Colorado
'132.163.4.102     time-b.timefreq.bldrdoc.gov    NIST, Boulder, Colorado
'132.163.4.103     time-c.timefreq.bldrdoc.gov    NIST, Boulder, Colorado
'192.43.244.18     time.nist.gov                  NCAR, Boulder, Colorado
'205.188.185.33    nist1.aol-va.truetime.com      AOL TrueTime, Virginia
'207.126.98.204    nist1-sj.glassey.com           Abovenet, San Jose, CA
'207.200.81.113    nist1.aol-ca.truetime.com      AOL TT, Sunnyvale, CA
'208.184.49.9      nist1-ny.glassey.com           Abovenet, New York City
'216.200.93.8      nist1-dc.glassey.com           Abovenet, Virginia
'66.243.43.21      nist1.datum.com                Datum, San Jose, CA
'132.163.4.102 anda bien


Private Const MAXIMADIFERENCIA = 8000 'diferencia máxima tolerable atribuída al tráfico en la red en ms

'Private Declare Function GetTickCount Lib "kernel32" () As Long

Private ZERO As Long
Private ULTIMO As Long
Public Advertencias As Integer
'Funciones del cliente
Public Sub GetNistTime(Optional ip As String = "132.163.4.102")
On Error GoTo ea
With frmMain
    If .WSAntiSH.State <> sckClosed Then
        .WSAntiSH.Close
        DoEvents
    End If
    .WSAntiSH.RemoteHost = ip
    .WSAntiSH.RemotePort = 13
    .WSAntiSH.Protocol = sckTCPProtocol
    .WSAntiSH.Connect
End With
ea:
End Sub
Public Sub AddTime(Tiempo As Long)
If ZERO = 0 Then
    ZERO = Tiempo - GetTickCount
Else
    ULTIMO = Tiempo - GetTickCount
    SendANTISH
End If
End Sub

Private Sub SendANTISH()
If (ULTIMO - ZERO) > MAXIMADIFERENCIA Or (ULTIMO - ZERO) < (0 - MAXIMADIFERENCIA) Then
Advertencias = Advertencias + 1

If Advertencias >= 5 Then End
'Call AddtoRichTextBox(frmMain.RecTxt, "SH+" & ULTIMO - ZERO, 255, 0, 0, True, False, False)
End If
End Sub

