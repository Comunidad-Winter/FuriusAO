Attribute VB_Name = "MD5"
'FIXIT: Utilice Option Explicit para evitar la creaci�n impl�cita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal T As Long, ByVal r As String)
Public Function MD5String(P As String) As String
Dim r As String * 32, T As Long

r = Space$(32)
T = Len(P)
MDStringFix P, T, r
MD5String = r
    
End Function
Public Function MD5File(f As String) As String
Dim r As String * 32

r = Space$(32)
MDFile f, r
MD5File = r

End Function

