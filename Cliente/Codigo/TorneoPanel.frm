VERSION 5.00
Begin VB.Form FrmTorneo 
   BorderStyle     =   0  'None
   Caption         =   "Quieren participar"
   ClientHeight    =   3570
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3150
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Menu Sumonear 
      Caption         =   "Sumonear"
      Index           =   0
   End
End
Attribute VB_Name = "FrmTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Sumonear_Click(Index As Integer)
Call SendData("/SUM " & List1.List(List1.ListIndex))
End Sub

