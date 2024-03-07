Attribute VB_Name = "MoD_MIDI"

Option Explicit

Public Const MIdi_Inicio = 6

Public CurMidi As String
Public LoopMidi As Byte
Public IsPlayingCheck As Boolean

Public GetStartTime As Long
Public Offset As Long
Public mtTime As Long
Public mtLength As Double
Public dTempo As Double


Dim timesig As DMUS_TIMESIGNATURE
Dim portcaps As DMUS_PORTCAPS

Dim msg As String
Dim time As Double
Dim Offset2 As Long
Dim ElapsedTime2 As Double
Dim fIsPaused As Boolean


Public Sub CargarMIDI(Archivo As String)

If Musica = 1 Then Exit Sub

On Error GoTo fin
    
    If IsPlayingCheck Then Stop_Midi
    If Loader Is Nothing Then Set Loader = DirectX.DirectMusicLoaderCreate()
    Set Seg = Loader.LoadSegment(Archivo)
        
   
        
    Set Loader = Nothing
    
    
    
    Exit Sub
fin:
    LogError "Error producido en "

End Sub
Public Sub Stop_Midi()

If IsPlayingCheck Then
     IsPlayingCheck = False
     Seg.SetStartPoint (0)
     Call Perf.Stop(Seg, SegState, 0, 0)
     
     Call Perf.Reset(0)
End If

End Sub

Public Sub Play_Midi()
If Musica = 1 Then Exit Sub
On Error GoTo fin
        
        
    
    Set SegState = Perf.PlaySegment(Seg, 0, 0)
    
    IsPlayingCheck = True
    Exit Sub
fin:
    LogError "Error producido en Public Sub Play_Midi()"

End Sub




