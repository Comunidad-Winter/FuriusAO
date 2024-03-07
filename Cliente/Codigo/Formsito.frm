VERSION 5.00
Begin VB.Form Formsito 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Yomando 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2880
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Formsito.frx":0000
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Formsito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Yomando = "" Then Exit Sub
SendData "~GLOBAL" & UserName & "> " & Yomando
Yomando = ""

End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

