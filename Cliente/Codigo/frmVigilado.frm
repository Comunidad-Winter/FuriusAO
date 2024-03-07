VERSION 5.00
Begin VB.Form frmVigilado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Panel Vigilados"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "VerPC - Procesos"
      End
      Begin VB.Menu mnuFPS 
         Caption         =   "VerFPS "
      End
      Begin VB.Menu mnuNovig 
         Caption         =   "Dejar de vigilar"
      End
   End
End
Attribute VB_Name = "frmVigilado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu menU_usuario
End If
End Sub


Private Sub mnuIR_Click()
Call SendData("/IRA " & List1)
End Sub

Private Sub mnutraer_Click()
Call SendData("/VERPC " & List1)
End Sub


Private Sub mnufps_Click()
Call SendData("/fps " & List1)
End Sub

Private Sub mnunovig_Click()
Call SendData("/novigilar " & List1)
If List1.ListIndex <> -1 Then
List1.RemoveItem List1.ListIndex
End If
End Sub

