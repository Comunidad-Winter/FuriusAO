Attribute VB_Name = "MD5"

Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal R As String)
Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal T As Long, ByVal R As String)

Public Function MD5String(p As String) As String

    Dim R As String * 32, T As Long
    R = Space(32)
    T = Len(p)
    MDStringFix p, T, R
    MD5String = R
End Function

Public Function MD5File(f As String) As String

    Dim R As String * 32
    R = Space(32)
    MDFile f, R
    MD5File = R
End Function



