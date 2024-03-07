VERSION 5.00
Begin VB.Form leo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de control "
   ClientHeight    =   4620
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   4920
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   7935
   End
   Begin VB.Label Label4 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   12120
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label ip 
      Caption         =   "Label1"
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
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   4350
      Width           =   3015
   End
   Begin VB.Label nombre 
      Caption         =   "nombre"
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
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   4350
      Width           =   4695
   End
   Begin VB.Line Line2 
      X1              =   4800
      X2              =   4800
      Y1              =   -360
      Y2              =   4680
   End
   Begin VB.Label Label3 
      Caption         =   "IP:"
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
      Left            =   5640
      TabIndex        =   5
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Nick 
      Caption         =   "Nick:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Aplicaciones 
      Caption         =   "Aplicaciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Procesos y ruta de ejecucion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu mnuPrc 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnucer 
         Caption         =   "Cerrar "
      End
      Begin VB.Menu mnuinf 
         Caption         =   "Informacion"
      End
      Begin VB.Menu mnutime 
         Caption         =   "Tiempo de ejecucion"
      End
   End
End
Attribute VB_Name = "leo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()
If MsgBox("¿Está seguro de cerrar el proceso seleccionado al usuario?", vbYesNo) = vbYes Then
Call SendData("/CERRARPROCESO " & nombre & "@" & ReadField(2, List2.List(List2.ListIndex), Asc(":")))
If List2.ListIndex <> -1 Then
   'Eliminamos el elemento que se encuentra seleccionado
   List2.RemoveItem List2.ListIndex
End If
End If
End Sub

Private Sub List1_Click()
If Button = vbRightButton Then
    PopupMenu mnuPrc
End If

End Sub

