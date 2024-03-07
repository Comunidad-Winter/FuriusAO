VERSION 5.00
Begin VB.Form Frmpenas 
   Caption         =   "Historial de penas v1.0"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "N°"
      Top             =   4965
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BORRAR PENA"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese aqui el numero de pena a borrar."
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   5040
      Width           =   3495
   End
End
Attribute VB_Name = "Frmpenas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Call senddata "(/BORRARPENA  )"
End Sub

