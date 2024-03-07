VERSION 5.00
Begin VB.Form TorneoPanel3vs3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Torneo FúriusAo  Modalidad: 3 Vs 3"
   ClientHeight    =   6345
   ClientLeft      =   3345
   ClientTop       =   615
   ClientWidth     =   9810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   3  'Dash-Dot
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6345
   ScaleMode       =   0  'User
   ScaleWidth      =   9810
   Begin VB.CommandButton Command29 
      Caption         =   "CERRAR"
      Height          =   255
      Left            =   8400
      TabIndex        =   69
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pasan"
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   8400
      TabIndex        =   68
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   8400
      TabIndex        =   67
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Cuenta"
      Height          =   375
      Left            =   600
      TabIndex        =   64
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   63
      Text            =   "10"
      Top             =   5940
      Width           =   375
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   8400
      TabIndex        =   58
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   8400
      TabIndex        =   57
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   8400
      TabIndex        =   56
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   8400
      TabIndex        =   55
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Jugadores Ronda Nº2"
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   0
      TabIndex        =   31
      Top             =   3120
      Width           =   8175
      Begin VB.CommandButton Command21 
         Caption         =   "Pasan "
         Height          =   375
         Left            =   6720
         TabIndex        =   54
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Pierden"
         Height          =   375
         Left            =   7320
         TabIndex        =   53
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   6960
         TabIndex        =   41
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Añadir"
         Height          =   255
         Left            =   5760
         TabIndex        =   40
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   1320
         TabIndex        =   39
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Añadir"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         TabIndex        =   37
         Text            =   "Jugador1-Jugador2-Jugador3"
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Text            =   "Jugador1-Jugador2-Jugador3"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.ListBox List3 
         Height          =   1425
         Left            =   4080
         TabIndex        =   35
         Top             =   480
         Width           =   3975
      End
      Begin VB.ListBox List4 
         Height          =   1425
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Pasan "
         Height          =   375
         Left            =   2280
         TabIndex        =   33
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Pierden"
         Height          =   375
         Left            =   2880
         TabIndex        =   32
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   66
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label25 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   65
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label12 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5760
         TabIndex        =   52
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   51
         Top             =   270
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "Total de equipos:"
         Height          =   255
         Left            =   4440
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Total de equipos:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   48
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   47
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   45
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   44
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   43
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "Vs.  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   42
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Pasan"
      Height          =   375
      Left            =   5400
      TabIndex        =   25
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Anunciar Pelea"
      Height          =   375
      Left            =   8400
      TabIndex        =   24
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Llamar Equipos"
      Height          =   375
      Left            =   8400
      TabIndex        =   23
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Llamar Equipos"
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Anunciar Pelea"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jugadores Ronda Nº1"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      Begin VB.CommandButton Command16 
         Caption         =   "Pierden"
         Height          =   375
         Left            =   7320
         TabIndex        =   30
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pierden"
         Height          =   375
         Left            =   2880
         TabIndex        =   29
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Pasan "
         Height          =   375
         Left            =   2280
         TabIndex        =   26
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   6960
         TabIndex        =   18
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Añadir"
         Height          =   255
         Left            =   5760
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Añadir"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         TabIndex        =   14
         Text            =   "Jugador1-Jugador2-Jugador3"
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "Jugador1-Jugador2-Jugador3"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.ListBox List2 
         Height          =   1425
         Left            =   4080
         TabIndex        =   2
         Top             =   480
         Width           =   3975
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label24 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   28
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label23 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   27
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label11 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000B&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   270
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Total de equipos:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Vs.  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "Total de equipos:"
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Zona A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   8280
      TabIndex        =   59
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "Zona B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   8280
      TabIndex        =   60
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      Caption         =   "Jugadores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   8280
      TabIndex        =   61
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      Caption         =   "Jugadores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   8280
      TabIndex        =   62
      Top             =   240
      Width           =   1455
   End
   Begin VB.Menu msg 
      Caption         =   "&Mensajes"
      Begin VB.Menu msg1 
         Caption         =   "¡Una extraordinaria pelea, ninguno de los equipos logra sacarse ventaja!"
         Shortcut        =   ^A
      End
      Begin VB.Menu msg2 
         Caption         =   "¡Esta pelea deja de que hablar, ambos equipos están dando un gran espectáculo!"
         Shortcut        =   ^B
      End
      Begin VB.Menu msg3 
         Caption         =   "Una pelea digna de ver, unos de los mejores enfrentamientos de este evento…"
         Shortcut        =   ^C
      End
      Begin VB.Menu msg4 
         Caption         =   "Remos, inmos, y mas inmos… ¡Que pelea muchachos!"
         Shortcut        =   ^D
      End
      Begin VB.Menu msg5 
         Caption         =   "¡Esquinas! ¡Mucha Suerte! Comienza en...    "
         Shortcut        =   ^E
      End
      Begin VB.Menu msg6 
         Caption         =   "Inmos, remos, apocas, una gran batalla, aun no se ha visto lo mejor!"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu bichos 
      Caption         =   "&NPC's"
      Begin VB.Menu bove 
         Caption         =   "Bóveda"
      End
      Begin VB.Menu sacer 
         Caption         =   "Sacerdote"
      End
      Begin VB.Menu potas 
         Caption         =   "Pociones"
      End
      Begin VB.Menu morfi 
         Caption         =   "Comida"
      End
      Begin VB.Menu agua 
         Caption         =   "Agua"
      End
   End
   Begin VB.Menu mod 
      Caption         =   "&Modalidad"
      Begin VB.Menu dosvsdos 
         Caption         =   "Torneo 2Vs2"
      End
      Begin VB.Menu cuatrovscuatro 
         Caption         =   "Torneo 4vs4"
      End
   End
   Begin VB.Menu torneo 
      Caption         =   "&Torneo"
   End
End
Attribute VB_Name = "TorneoPanel3vs3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub agua_Click()
Call SendData("/ACC 137")
End Sub

Private Sub bove_Click()
Call SendData("/ACC 57")
End Sub

Private Sub Command1_Click()
If List3.ListIndex <> -1 Then
List3.RemoveItem List3.ListIndex
Label12 = List3.ListCount
End If
End Sub

Private Sub Command10_Click()
If List2.ListIndex <> -1 Then
List2.RemoveItem List2.ListIndex
Label11 = List2.ListCount

End If
End Sub

Private Sub Command11_Click()
If List4.ListIndex <> -1 Then
List4.RemoveItem List4.ListIndex
Label13 = List4.ListCount
End If
End Sub

Private Sub Command12_Click()
List4.AddItem Text2.Text
Label13 = List4.ListCount
End Sub

Private Sub Command13_Click()
If MsgBox("¿Está seguro que desea llamar a " & "[" & List3 & "]" & " Vs. " & "[" & List4 & "]" & " ?", vbYesNo) = vbYes Then
Call SendData("/SUM " & ReadField(1, List4, Asc("-")))
Call SendData("/SUM " & ReadField(2, List4, Asc("-")))
Call SendData("/SUM " & ReadField(3, List4, Asc("-")))
Call SendData("/SUM " & ReadField(1, List3, Asc("-")))
Call SendData("/SUM " & ReadField(2, List3, Asc("-")))
Call SendData("/SUM " & ReadField(3, List3, Asc("-")))
End If
End Sub

Private Sub Command14_Click()
If MsgBox("¿Está seguro que desea anunciar la pelea seleccionada?", vbYesNo) = vbYes Then
Call SendData("/rmsg " & "Juegan la siguiente pelea: " & List4 & " Vs. " & List3)
End If
End Sub

Private Sub Command16_Click()
If MsgBox("¿Está seguro que desea que " & "[" & List2 & "]" & " pierdan esta ronda ronda?", vbYesNo) = vbYes Then
Call SendData("/EXPLOTA " & ReadField(1, List2, Asc("-")))
Call SendData("/EXPLOTA " & ReadField(2, List2, Asc("-")))
Call SendData("/EXPLOTA " & ReadField(3, List2, Asc("-")))
Call SendData("/rmsg " & "Quedan eliminados del torneo: " & List2)
List2.RemoveItem List2.ListIndex
Label11 = List2.ListCount
End If
End Sub

Private Sub Command17_Click()
If MsgBox("¿Está seguro que desea que " & "[" & List4 & "]" & " pasan de ronda", vbYesNo) = vbYes Then
If List4.ListIndex <> -1 Then
List1.AddItem List4
Call SendData("/PASAN " & ReadField(1, List4, Asc("-")))
Call SendData("/PASAN " & ReadField(2, List4, Asc("-")))
Call SendData("/PASAN " & ReadField(3, List4, Asc("-")))
Call SendData("/rmsg " & "Pasan a la siguiente instancia: " & List4)
List4.RemoveItem List4.ListIndex
Label10 = List1.ListCount
Label13 = List4.ListCount
End If
End If
End Sub

Private Sub Command18_Click()
If MsgBox("¿Está seguro que desea que " & "[" & List1 & "]" & " pasan de ronda", vbYesNo) = vbYes Then
If List1.ListIndex <> -1 Then
List4.AddItem List1
Call SendData("/PASAN " & ReadField(1, List1, Asc("-")))
Call SendData("/PASAN " & ReadField(2, List1, Asc("-")))
Call SendData("/PASAN " & ReadField(3, List1, Asc("-")))
Call SendData("/rmsg " & "Pasan a la siguiente instancia: " & List1)
List1.RemoveItem List1.ListIndex
Label10 = List1.ListCount
Label13 = List4.ListCount
End If
End If
End Sub

Private Sub Command19_Click()
If MsgBox("¿Está seguro que desea que " & "[" & List4 & "]" & " pierdan esta ronda ronda?", vbYesNo) = vbYes Then
Call SendData("/EXPLOTA " & ReadField(1, List4, Asc("-")))
Call SendData("/EXPLOTA " & ReadField(2, List4, Asc("-")))
Call SendData("/EXPLOTA " & ReadField(3, List4, Asc("-")))
Call SendData("/rmsg " & "Quedan eliminados del torneo: " & List4)
List4.RemoveItem List4.ListIndex
Label13 = List4.ListCount
End If
End Sub

Private Sub Command2_Click()
List3.AddItem Text1.Text
Label12 = List3.ListCount
End Sub

Private Sub Command20_Click()
If MsgBox("¿Está seguro que desea que " & "[" & List3 & "]" & " pierdan esta ronda ronda?", vbYesNo) = vbYes Then
Call SendData("/EXPLOTA " & ReadField(1, List3, Asc("-")))
Call SendData("/EXPLOTA " & ReadField(2, List3, Asc("-")))
Call SendData("/EXPLOTA " & ReadField(3, List3, Asc("-")))
Call SendData("/rmsg " & "Quedan eliminados del torneo: " & List3)
List3.RemoveItem List3.ListIndex
Label12 = List3.ListCount
End If
End Sub

Private Sub Command21_Click()
If MsgBox("¿Está seguro que desea que " & "[" & List3 & "]" & " pasan de ronda", vbYesNo) = vbYes Then
If List3.ListIndex <> -1 Then
List2.AddItem List3
Call SendData("/PASAN " & ReadField(1, List3, Asc("-")))
Call SendData("/PASAN " & ReadField(2, List3, Asc("-")))
Call SendData("/PASAN " & ReadField(3, List3, Asc("-")))
Call SendData("/rmsg " & "Pasan a la siguiente instancia: " & List3)
List3.RemoveItem List3.ListIndex
Label11 = List2.ListCount
Label12 = List3.ListCount
End If
End If
End Sub

Private Sub Command22_Click()
If MsgBox("¿Está seguro que desea Guardar la Zona Nº1?", vbYesNo) = vbYes Then
Call GuardarLista(List1, "C:/list1.txt")
Call GuardarLista(List2, "C:/list2.txt")
End If
End Sub

Private Sub Command23_Click()
If MsgBox("¿Está seguro que desea cargar la Zona Nº1?", vbYesNo) = vbYes Then
Call LeerLista(List1, "C:/list1.txt")
Call LeerLista(List2, "C:/list2.txt")
End If
End Sub

Private Sub Command24_Click()
If MsgBox("¿Está seguro que desea guardar la Zona Nº2?", vbYesNo) = vbYes Then
Call GuardarLista(List4, "C:/list4.txt")
Call GuardarLista(List3, "C:/list3.txt")
End If
End Sub

Private Sub Command25_Click()
If MsgBox("¿Está seguro que desea cargar la Zona Nº2?", vbYesNo) = vbYes Then
Call LeerLista(List4, "C:/list4.txt")
Call LeerLista(List3, "C:/list3.txt")
End If
End Sub

Private Sub Command26_Click()
Call SendData("/cuenta " & Text5.Text)
End Sub

Private Sub Command27_Click()
If MsgBox("¿Está seguro que desea borrar la Zona Nº2?", vbYesNo) = vbYes Then
List3.Clear
List4.Clear
Label13 = List3.ListCount
Label12 = List4.ListCount
End If
End Sub

Private Sub Command28_Click()
If MsgBox("¿Está seguro que desea borrar toda la Zona Nº1?", vbYesNo) = vbYes Then
List1.Clear
List2.Clear
Label10 = List1.ListCount
Label11 = List2.ListCount
End If
End Sub

Private Sub Command29_Click()
Call GuardarLista(List1, "C:/list1.txt")
Call GuardarLista(List2, "C:/list2.txt")
Call GuardarLista(List3, "C:/list3.txt")
Call GuardarLista(List4, "C:/list4.txt")
End Sub

Private Sub Command3_Click()
If MsgBox("¿Está seguro que desea anunciar la pelea seleccionada?", vbYesNo) = vbYes Then
Call SendData("/rmsg " & "Juegan la siguiente pelea: " & List1 & " Vs. " & List2)
End If
End Sub

Private Sub Command4_Click()
If MsgBox("¿Está seguro que desea llamar a " & "[" & List1 & "]" & " Vs. " & "[" & List2 & "]" & " ?", vbYesNo) = vbYes Then
Call SendData("/SUM " & ReadField(1, List1, Asc("-")))
Call SendData("/SUM " & ReadField(2, List1, Asc("-")))
Call SendData("/SUM " & ReadField(3, List1, Asc("-")))
Call SendData("/SUM " & ReadField(1, List2, Asc("-")))
Call SendData("/SUM " & ReadField(2, List2, Asc("-")))
Call SendData("/SUM " & ReadField(3, List2, Asc("-")))
End If
End Sub

Private Sub Command5_Click()
If MsgBox("¿Está seguro que desea que " & "[" & List2 & "]" & " pasan de ronda", vbYesNo) = vbYes Then
If List2.ListIndex <> -1 Then
List3.AddItem List2
Call SendData("/PASAN " & ReadField(1, List2, Asc("-")))
Call SendData("/PASAN " & ReadField(2, List2, Asc("-")))
Call SendData("/PASAN " & ReadField(3, List2, Asc("-")))
Call SendData("/rmsg " & "Pasan a la siguiente instancia: " & List2)
List2.RemoveItem List2.ListIndex
Label11 = List2.ListCount
Label12 = List3.ListCount
End If
End If
End Sub

Private Sub Command6_Click()
If MsgBox("¿Está seguro que desea que " & "[" & List1 & "]" & " pierdan esta ronda ronda?", vbYesNo) = vbYes Then
Call SendData("/EXPLOTA " & ReadField(1, List1, Asc("-")))
Call SendData("/EXPLOTA " & ReadField(2, List1, Asc("-")))
Call SendData("/EXPLOTA " & ReadField(3, List1, Asc("-")))
Call SendData("/rmsg " & "Quedan eliminados del torneo: " & List1)
List1.RemoveItem List1.ListIndex
Label10 = List1.ListCount
End If
End Sub

Private Sub Command7_Click()
List1.AddItem Text3.Text
Label10 = List1.ListCount
End Sub

Private Sub Command8_Click()
If List1.ListIndex <> -1 Then
List1.RemoveItem List1.ListIndex
Label10 = List1.ListCount

End If
End Sub

Private Sub Command9_Click()
List2.AddItem Text4.Text
Label11 = List2.ListCount
End Sub

Private Sub cuatrovscuatro_Click()
If MsgBox("¿Está seguro que desea cambiar a Modalidad 4vs4?", vbYesNo) = vbYes Then
Unload Me
TorneoPanel4vs4.Show
End If
End Sub

Private Sub dosvsdos_Click()
If MsgBox("¿Está seguro que desea cambiar a Modalidad 2vs2?", vbYesNo) = vbYes Then
Unload Me
TorneoPanel.Show
End If
End Sub

Private Sub Form_Load()
 
List1.DragIcon = LoadPicture(App.Path & "\Graficos\Drag.ico")
List2.DragIcon = LoadPicture(App.Path & "\Graficos\Drag.ico")
List3.DragIcon = LoadPicture(App.Path & "\Graficos\Drag.ico")
List4.DragIcon = LoadPicture(App.Path & "\Graficos\Drag.ico")

Call LeerLista(List1, "C:/list1.txt")
Call LeerLista(List2, "C:/list2.txt")
Call LeerLista(List3, "C:/list3.txt")
Call LeerLista(List4, "C:/list4.txt")
End Sub



Private Sub Label23_Click()
If List1.ListIndex <> -1 Then
List2.AddItem List1
List1.RemoveItem List1.ListIndex
Label10 = List1.ListCount
Label11 = List2.ListCount
End If
End Sub

Private Sub Label24_Click()
If List2.ListIndex <> -1 Then
   'Eliminamos el elemento que se encuentra seleccionado
   List1.AddItem List2
List2.RemoveItem List2.ListIndex
Label11 = List2.ListCount
Label10 = List1.ListCount

End If
End Sub

Private Sub Label25_Click()
If List3.ListIndex <> -1 Then
List4.AddItem List3
List3.RemoveItem List3.ListIndex
Label13 = List4.ListCount
Label12 = List3.ListCount
End If
End Sub

Private Sub Label26_Click()
If List4.ListIndex <> -1 Then
List3.AddItem List4
List4.RemoveItem List4.ListIndex
Label13 = List4.ListCount
Label12 = List3.ListCount
End If
End Sub

Private Sub List2_Click()
    SincListBox List2, List1
 '  Label11 = List2.ListCount
End Sub

Private Sub List1_Scroll()
    'Sincronizar también el primer item mostrado en la lista
   List2.TopIndex = List1.TopIndex
End Sub
Private Sub List2_Scroll()
    'Sincronizar también el primer item mostrado en la lista
    List1.TopIndex = List2.TopIndex
End Sub

Private Sub List1_Click()
    SincListBox List1, List2
   ' Label10 = List1.ListCount
End Sub
Private Sub QuitarListSelected(unList As Control)
    'Quitar los elementos seleccionados del listbox indicado
    'Parámetros:
    '   unList      el List a controlar
    '
    Dim i&
    
    With unList
        'Sólo hacer el bucle si permite multiselección
        If .MultiSelect Then
            For i = 0 To .ListCount - 1
                .Selected(i) = False
            Next
        End If
    End With
End Sub

Private Sub ListSelected(elListOrig As Control, elListDest As Control)
    'Marca en el ListDest los elementos seleccionados del ListOrig
    '
    'Los dos listbox deben tener el mismo número de elementos
    '
    Dim i&
    
    'Por si no tienen los mismos elementos
    On Local Error Resume Next
    
    With elListOrig
        For i = 0 To .ListCount - 1
            'Si el origen está seleccionado...
            If .Selected(i) Then
                elListDest.Selected(i) = .Selected(i)
            Else
                'sino, quitar la posible selección
                elListDest.Selected(i) = False
            End If
        Next
    End With
        
    Err = 0
End Sub

Private Sub PonerListSelected(elListOrig As Control, elListDest As Control)
    'Marca en el ListDest los elementos seleccionados del ListOrig
    '
    'Los dos listbox deben tener el mismo número de elementos
    '
    Dim i&
    
    'Por si no tienen los mismos elementos
    On Local Error Resume Next
    
    With elListOrig
        For i = 0 To .ListCount - 1
            elListDest.Selected(i) = .Selected(i)
        Next
    End With
        
    Err = 0
End Sub

Private Sub SincListBox(elListOrig As Control, elListDest As Control)
    Static EnListBox As Boolean
        
    'Sincronizar el elListDest con el elListOrig
    If Not EnListBox Then
    
        EnListBox = True
        
'        'Desmarcar los elementos seleccionados
'        QuitarListSelected elListDest
'
'        'Marcar en el 1º ListBox los seleccionados del 2º
'        PonerListSelected elListOrig, elListDest
        
        'Poner en el ListDest los mismos que en ListOrig
        ListSelected elListOrig, elListDest
        
        'Posicionar el elemento superior
     '   elListDest.TopIndex = elListOrig.TopIndex
        
        EnListBox = False
    End If
End Sub

Private Sub List3_Click()
SincListBox List3, List4
'Label12 = List3.ListCount
End Sub

Private Sub List4_Click()
SincListBox List4, List3
'Label13 = List4.ListCount
End Sub

Sub GuardarLista(listax As ListBox, Donde As String)
Dim fnum As Integer
On Error GoTo Ninguno
    fnum = FreeFile
    Open Donde For Output As fnum
    
    Dim i As Integer
    For i = 0 To listax.ListCount
            Print #fnum, listax.List(i)
        DoEvents
    Next i
    
    Close fnum
    MsgBox "Torneo Guardado."
Ninguno:
End Sub

Sub LeerLista(listax As ListBox, Donde As String)
Dim fnum As Integer
Dim Txt As String
On Error GoTo Ninguno

fnum = FreeFile
    Open Donde For Input As fnum
    Do While Not EOF(fnum)
        Line Input #fnum, Txt
        listax.AddItem Txt
        'Texto.Text = Texto.Text & vbCrLf & txt
    Loop
    Close fnum
    MsgBox "Torneo Cargado."
    Label10 = List1.ListCount
    Label11 = List2.ListCount
    Label13 = List4.ListCount
    Label12 = List3.ListCount
Ninguno:
End Sub

Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
    
    ' Si el control es el List2 entonces OK ..
    If Source Is List4 Then
       If List4.ListIndex <> -1 Then
            List1.AddItem List4.List(List4.ListIndex)
            List4.RemoveItem List4.ListIndex
            Label10 = List1.ListCount
            Label13 = List4.ListCount
       End If
    End If
End Sub


' Inicia la operación de arrastre, es decir el drag para List1
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, _
                            X As Single, Y As Single)
    List1.Drag vbBeginDrag
End Sub

' Al soltar el item en List2, se agrega al mismo y se elimina el del List1

Private Sub List4_DragDrop(Source As Control, _
                           X As Single, Y As Single)
    
    ' Si el control es el List1 entonces..
    If Source Is List1 Then
        If List1.ListIndex <> -1 Then
           List4.AddItem List1.List(List1.ListIndex)
           List1.RemoveItem List1.ListIndex
           Label10 = List1.ListCount
           Label13 = List4.ListCount
        End If
    End If
End Sub

' Comienza el Drag para el List2
Private Sub List4_MouseDown(Button As Integer, Shift As Integer, _
                                        X As Single, Y As Single)
    List4.Drag vbBeginDrag
End Sub

'#####################################################################
'####################### AHORA LIST2 y LIST3 #########################
'#####################################################################

Private Sub List2_DragDrop(Source As Control, X As Single, Y As Single)
    
    ' Si el control es el List2 entonces OK ..
    If Source Is List3 Then
       If List3.ListIndex <> -1 Then
            List2.AddItem List3.List(List3.ListIndex)
            List3.RemoveItem List3.ListIndex
            Label11 = List2.ListCount
            Label12 = List3.ListCount
       End If
    End If
End Sub


' Inicia la operación de arrastre, es decir el drag para List1
Private Sub List2_MouseDown(Button As Integer, Shift As Integer, _
                            X As Single, Y As Single)
    List2.Drag vbBeginDrag
End Sub

' Al soltar el item en List2, se agrega al mismo y se elimina el del List1

Private Sub List3_DragDrop(Source As Control, _
                           X As Single, Y As Single)
    
    ' Si el control es el List1 entonces..
    If Source Is List2 Then
        If List2.ListIndex <> -1 Then
           List3.AddItem List2.List(List2.ListIndex)
           List2.RemoveItem List2.ListIndex
           Label11 = List2.ListCount
           Label12 = List3.ListCount
        End If
    End If
End Sub

' Comienza el Drag para el List2
Private Sub List3_MouseDown(Button As Integer, Shift As Integer, _
                                        X As Single, Y As Single)
    List3.Drag vbBeginDrag
End Sub

Private Sub morfi_Click()
Call SendData("/ACC 134")
End Sub

Private Sub msg1_Click()
If MsgBox("¿Está seguro que desea mandar el mensaje?", vbYesNo) = vbYes Then
Call SendData("/rmsg " & "¡Una extraordinaria pelea, ninguno de los equipos logra sacarse ventaja!")
End If
End Sub

Private Sub msg2_Click()
If MsgBox("¿Está seguro que desea mandar el mensaje?", vbYesNo) = vbYes Then
Call SendData("/rmsg " & "¡Esta pelea deja mucho que hablar, ambos equipos están dando un gran espectáculo!")
End If
End Sub

Private Sub msg3_Click()
If MsgBox("¿Está seguro que desea mandar el mensaje?", vbYesNo) = vbYes Then
Call SendData("/rmsg " & "Una pelea digna de ver, unos de los mejores enfrentamientos de este evento…")
End If
End Sub

Private Sub msg4_Click()
If MsgBox("¿Está seguro que desea mandar el mensaje?", vbYesNo) = vbYes Then
Call SendData("/rmsg " & "Remos, inmos, y mas inmos… ¡Que pelea muchachos!")
End If
End Sub

Private Sub msg5_Click()
If MsgBox("¿Está seguro que desea mandar el mensaje?", vbYesNo) = vbYes Then
Call SendData("/rmsg " & "¡Esquinas! ¡Mucha Suerte! Comienza en...")
End If
End Sub

Private Sub msg6_Click()
If MsgBox("¿Está seguro que desea mandar el mensaje?", vbYesNo) = vbYes Then
Call SendData("/rmsg " & "Inmos, remos, apocas, una gran batalla, aun no se ha visto lo mejor!")
End If
End Sub

Private Sub potas_Click()
Call SendData("/ACC 14")
End Sub

Private Sub sacer_Click()
Call SendData("/ACC 5")
End Sub

Private Sub torneo_Click()
If MsgBox("¿Está seguro que desea abrir el torneo y avisar del mismo?", vbYesNo) = vbYes Then
Call SendData("/TORNEO")
Call SendData("/rmsg " & "Ya se encuentran abiertas las inscripciones para poder participar deberás enviar /participar. ¡Muchas Gracias!")
End If
End Sub
