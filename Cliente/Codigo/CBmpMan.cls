VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBmpMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Enum eModoBMPMan
    mododinamico = 0
    ModoCacheado = 1
End Enum

Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type BMP_DB_ENTRY
    FileName As Integer
    Height As Long
    Width As Long
    Size As Long
    Accesos As Long
    Efecto As Long
    Cont As Long
End Type

Private Type BMP_INFO
    Height As Long
    Width As Long
    Size As Long
End Type

Private mModo As eModoBMPMan
Private mBMPDB() As BMP_DB_ENTRY
Private mDDraw As DirectDraw7
Private mSurfaceDB() As DirectDrawSurface7
Private mKeys() As Integer

Private mMaxEntries As Long
Private mContador As Long
Private mCantidadGraficos As Long

Private mEfectoPred As Integer

Private GrhPath As String

Public Property Let Path(ByVal inString As String)
GrhPath = inString
End Property
Public Function GetBMP(FileName As Integer, Optional ByVal Efecto As Integer = -1) As DirectDrawSurface7
Dim RetVal As Integer

If Efecto < 0 Then
    Efecto = mEfectoPred
End If



'bNoche = Not bNoche
'SurfaceDB.EfectoPred = IIf(bNoche, 1, 0)
'SurfaceDB.BorrarTodo

Select Case mModo
Case mododinamico
    If mKeys(Efecto, FileName) <> 0 Then
        mBMPDB(mKeys(Efecto, FileName)).Accesos = mBMPDB(mKeys(Efecto, FileName)).Accesos + 1
        Set GetBMP = mSurfaceDB(mKeys(Efecto, FileName))
    Else
            RetVal = BuscaVacio()
            If RetVal = -1 Then
            RetVal = 1
            'RetVal = BorraMenosUsado()
            RetVal = BorrarTodo
        End If
        Call CrearGrafico(FileName, RetVal)
        mBMPDB(RetVal).Accesos = mBMPDB(RetVal).Accesos + 1
        mBMPDB(RetVal).Efecto = Efecto
        mKeys(Efecto, FileName) = RetVal
        Set GetBMP = mSurfaceDB(RetVal)

  '  #If ConAlfaB = 1 Then
   ' If Efecto = 1 Then
   '     EfectoNoche mSurfaceDB(RetVal)
   ' End If
   ' #End If

    
    
    End If
    
    
    

Case ModoCacheado
    Set GetBMP = mSurfaceDB(FileName)
End Select

End Function
Public Function Init(Optional modo As eModoBMPMan = ModoCacheado) As Long
On Error Resume Next
Dim loopc As Integer
mModo = modo

Select Case modo
    Case mododinamico
        If mMaxEntries < 1 Then
            Init = -1
            Exit Function
        End If
        
        ReDim mBMPDB(0 To mMaxEntries)
        ReDim mSurfaceDB(0 To mMaxEntries)
        ReDim mKeys(0 To 1, 0 To 17000) As Integer
        mEfectoPred = 0
        mCantidadGraficos = 0
        mContador = 0
        Init = 1
    
    Case ModoCacheado
        ReDim mSurfaceDB(1 To Config_Inicio.NumeroDeBMPs)
        For loopc = 1 To Config_Inicio.NumeroDeBMPs
            Call CrearGrafico(loopc, loopc)
        Next
    
End Select

Exit Function

End Function

'Public Function BorrarTodo2() As Long

'On Error Resume Nex't'
'Dim i As Integer '

'Select Case mModo
'Case mododinamico
'    For i = 1 To mMaxEntries
 '       Call ZeroMemory(mBMPDB(i), Len(mBMPDB(i)))
'        Set mSurfaceDB(i) = Nothing
'    Next
'Case ModoCacheado
'    For i = 1 To Config_Inicio.NumeroDeBMPs
'        Call ZeroMemory(mBMPDB(i), Len(mBMPDB(i)))
'        Set mSurfaceDB(i) = Nothing
'    Next
'End Select

'BorrarTodo = 1
'mCantidadGraficos = 0
'End Function

Public Function BorrarTodo() As Long
'Vacia la lista de graficos
Dim i As Integer
For i = 1 To mMaxEntries
    If mBMPDB(i).FileName > 0 Then
    Call Borrar(mBMPDB(i).FileName, mBMPDB(i).Efecto)
    End If
Next
BorrarTodo = 1
mCantidadGraficos = 0
End Function


Public Function Borrar(ByVal FileName As Integer, ByVal Efecto As Integer) As Integer
'Borra un grafico
Borrar = -1
Call ZeroMemory(mBMPDB(mKeys(Efecto, FileName)), Len(mBMPDB(mKeys(Efecto, FileName))))
Set mSurfaceDB(mKeys(Efecto, FileName)) = Nothing
mKeys(Efecto, FileName) = 0
Borrar = 1
mCantidadGraficos = mCantidadGraficos - 1
End Function

'Public Function Borrar2(FileName As Integer) As Integer

'Borrar = -1
'Call ZeroMemory(mBMPDB(mKeys(FileName)), Len(mBMPDB(mKeys(FileName))))
'Set mSurfaceDB(mKeys(FileName)) = Nothing
'mKeys(FileName) = 0
'Borrar = 1
'mCantidadGraficos = mCantidadGraficos - 1
'End Function

Function EsArbol(Numero As Integer)
EsArbol = False
If Numero > 933 And Numero < 972 Then EsArbol = True: Exit Function
If Numero > 6999 And Numero < 7009 Then EsArbol = True: Exit Function
If Numero = 644 Or Numero = 645 Or Numero = 648 Then EsArbol = True: Exit Function
End Function
Private Function CrearGrafico(Archivo As Integer, bIndex As Integer) As Integer

Dim ddsd As DDSURFACEDESC2, BMPInfo As BMP_INFO, RetVal As Long, ddck As DDCOLORKEY
ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH

Call GetBMPHeader(GrhPath & Archivo & ".bmp", BMPInfo)

With ddsd
    If RenderMod.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    
    .lHeight = BMPInfo.Height
    .lWidth = BMPInfo.Width
End With

ddck.high = 0: ddck.low = 0

If mModo = mododinamico Then
    With mBMPDB(bIndex)
        .FileName = Archivo
        .Accesos = 0
        .Height = BMPInfo.Height
        .Width = BMPInfo.Width
        .Size = BMPInfo.Size
        .Cont = mContador
    End With
    mContador = mContador + 1

mCantidadGraficos = mCantidadGraficos + 1
End If
If Archivo <> 0 Then Set mSurfaceDB(bIndex) = mDDraw.CreateSurfaceFromFile(GrhPath & Archivo & ".bmp", ddsd)
mSurfaceDB(bIndex).SetColorKey DDCKEY_SRCBLT, ddck

CrearGrafico = 1

End Function

Private Sub Class_Initialize()
mMaxEntries = 0
mEfectoPred = 0
End Sub

Private Function BuscaVacio() As Integer
Dim i As Integer
For i = 1 To mMaxEntries
    If (mSurfaceDB(i) Is Nothing) Then
        BuscaVacio = i
        Exit Function
    End If
Next

BuscaVacio = -1

End Function



Private Function BorraMenosUsado() As Integer
BorraMenosUsado = 1
'
Dim lValor As Long
Dim lCont As Long '

lValor = mBMPDB(1).Accesos
lCont = mBMPDB(1).Cont '

Dim i As Integer

For i = 1 To mMaxEntries
    If mBMPDB(i).Accesos < lValor And mBMPDB(i).Cont < lCont Then
        lValor = mBMPDB(i).Accesos
        lCont = mBMPDB(i).Cont
        BorraMenosUsado = i
    End If
Next
mCantidadGraficos = mCantidadGraficos - 1 '

mKeys(mBMPDB(BorraMenosUsado).FileName) = 0
Set mSurfaceDB(BorraMenosUsado) = Nothing
Call ZeroMemory(mBMPDB(BorraMenosUsado), Len(mBMPDB(BorraMenosUsado))) '

End Function

Public Property Let lpDirectDraw7(ByRef objDirectDraw7 As DirectDraw7)
Set mDDraw = objDirectDraw7
End Property
Private Function GetBMPHeader(Archivo As String, ByRef lpBMPInfo As BMP_INFO) As Long
Dim nfile As Integer, H_BMPH As BITMAPFILEHEADER, H_BMPI As BITMAPINFOHEADER
nfile = FreeFile()

Open Archivo For Binary Access Read As #nfile
Get #nfile, , H_BMPH
Get #nfile, , H_BMPI
Close #nfile






lpBMPInfo.Height = H_BMPI.biHeight
lpBMPInfo.Width = H_BMPI.biWidth
lpBMPInfo.Size = H_BMPH.bfSize

GetBMPHeader = 1

End Function

Public Property Get MaxEntries() As Long
MaxEntries = mMaxEntries
End Property

Public Property Let MaxEntries(ByVal vNewValue As Long)
mMaxEntries = vNewValue
End Property

Public Property Get CantidadGraficos() As Variant
CantidadGraficos = mCantidadGraficos
End Property

Public Property Get EfectoPred() As Integer
EfectoPred = mEfectoPred
End Property

Public Property Let EfectoPred(ByVal iNewValue As Integer)
mEfectoPred = iNewValue
End Property

