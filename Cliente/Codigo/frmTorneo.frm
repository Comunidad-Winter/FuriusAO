VERSION 5.00
Begin VB.Form frmTorneo 
   BorderStyle     =   0  'None
   Caption         =   "Quieren participar"
   ClientHeight    =   3570
   ClientLeft      =   2760
   ClientTop       =   4005
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Left            =   2350
      TabIndex        =   3
      Top             =   15
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Jugadores:"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu Sumonear 
      Caption         =   "Sumonear"
      Index           =   0
   End
End
Attribute VB_Name = "frmTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload Me
End Sub



Private Sub Sumonear_Click(index As Integer)
If List1.ListIndex = -1 Then Exit Sub
Call SendData("/SUM " & ReadField(1, List1, Asc(":")))
List1.RemoveItem List1.ListIndex
Label2 = List1.ListCount
End Sub

Private Sub Form_Load()


'Label2 = List1.ListCount
End Sub
