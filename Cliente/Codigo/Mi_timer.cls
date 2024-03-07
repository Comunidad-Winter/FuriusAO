VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mi_timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Intervalo As Long
Private QueMiro As Long
Private Tiempo As Double
Public Puede As Boolean
'Const QueMiro = 1000 'es si me fijo en milisegundo o cualquier metodo que use
'Const intervalo = 1000 ' en el metodo que uno use
'en este caso miro los milisegundos
'y cada un 1000 milisegundos(un segundo) ejecuto la accion

Public Sub init(TQueMiro As Long, Inter As Long)
    Intervalo = Inter
    QueMiro = TQueMiro

End Sub

Public Function puedo() As Boolean
    If (Timer * Intervalo) < Intervalo - 1 Then 'me fijo que  el tiempo no vuelva a cerro
        Tiempo = (Timer * QueMiro) + Intervalo
    End If
    If (Timer * QueMiro) >= Tiempo Then
        Puede = True
        
        Tiempo = (Timer * QueMiro) + Intervalo
    End If

puedo = Puede
Puede = False


End Function









Private Sub Class_Terminate()

DoEvents
End Sub

