VERSION 5.00
Begin VB.Form frmSalir 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image cancelar 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1920
      MouseIcon       =   "frmSalir.frx":0000
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image aceptar 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   720
      MouseIcon       =   "frmSalir.frx":030A
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "frmSalir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Aceptar_Click()

Call SendData("/SALIR")
Pocho = True
frmMain.Visible = False
FormBarInv.Visible = False
FormConsola.Visible = False
FormInfo.Visible = False
FormInv.Visible = False
FormListOpciones.Visible = False
Unload Me
Unload frmMain

End Sub

Private Sub Cancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "Salir.gif")


End Sub
