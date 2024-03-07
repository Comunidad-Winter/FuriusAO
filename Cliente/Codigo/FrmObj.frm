VERSION 5.00
Begin VB.Form FrmObj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Objetos"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox FraseCompleta 
      Caption         =   "Frase Completa"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "1"
      Top             =   3210
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Numeros de objetos:"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   3000
      Width           =   735
   End
End
Attribute VB_Name = "FrmObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
' Declaración del api SendMessage
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal Hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As String) As Long
  
'Constante "Mensaje" para buscar una cadena en el ListBox
Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
  
'Recibe la cadena y el valor de tipo boolean para _
 determinar si busca o no la cadena completa
Private Sub Buscar_ListBox(Frase As String)
  
Dim indice As Long
       
    ' Tipo de búsqueda
    If FraseCompleta Then
       indice = SendMessage(List1.Hwnd, LB_FINDSTRINGEXACT, -1, Frase)
    Else
       indice = SendMessage(List1.Hwnd, LB_FINDSTRING, -1, Frase)
    End If
       
       
    If indice Then
        ' se encontró la frase entonces la selecciona
        List1.ListIndex = indice
    End If
End Sub
  

  


Private Sub Command1_Click()



Call SendData("/ITEM " & ReadField(2, List1, Asc("-")) & " " & Text1.Text)
End Sub



Private Sub Text2_Change()
Call Buscar_ListBox(Text2)
End Sub


