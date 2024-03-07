Attribute VB_Name = "ModHash"


Option Explicit
Public Function GenHash(FileName As String) As String

Dim cStream As New cBinaryFileStream
Dim cCRC32 As New cCRC32
Dim lCRC32 As Long

cStream.File = FileName
lCRC32 = cCRC32.GetFileCrc32(cStream)
GenHash = hex$(lCRC32)

End Function
