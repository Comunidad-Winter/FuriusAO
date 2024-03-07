Attribute VB_Name = "Mod_ErrorLOG"


Option Explicit

Public Sub LogError(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\errores.log" For Append As #nfile
Print #nfile, desc
Close #nfile
End Sub

