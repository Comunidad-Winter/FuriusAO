VERSION 5.00
Begin VB.Form frmBotanica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alquimia"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tCant 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "1"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      MouseIcon       =   "frmBotanica.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2550
      Width           =   1710
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Construir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2505
      MouseIcon       =   "frmBotanica.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2550
      Width           =   1710
   End
   Begin VB.ListBox lstArmas 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4080
   End
End
Attribute VB_Name = "frmBotanica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
On Error Resume Next
If Int(Val(tCant)) < 1 Or Int(Val(tCant)) > 1000 Then
    MsgBox "La cantidad es invalida.", vbCritical
    Exit Sub
End If
Call SendData("CNA" & ObjBotanica(lstArmas.ListIndex) & "," & tCant.Text)

Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub
