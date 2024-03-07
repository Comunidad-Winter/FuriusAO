VERSION 5.00
Begin VB.Form frmParty2 
   BorderStyle     =   0  'None
   Caption         =   "Party"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Rechazar 
      Height          =   255
      Left            =   1800
      MouseIcon       =   "Party2.frx":0000
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image Aceptar 
      Height          =   255
      Left            =   600
      MouseIcon       =   "Party2.frx":030A
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Juancito"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   795
      Width           =   975
   End
End
Attribute VB_Name = "frmParty2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Acepta_Click(Index As Integer)

End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "invitacionparty.gif")
End Sub

Private Sub Rechazar_Click()
Call SendData("PARREC")
frmParty2.Visible = False
End Sub
Private Sub Aceptar_Click()
Call SendData("PARACE")
frmParty2.Visible = False
End Sub
