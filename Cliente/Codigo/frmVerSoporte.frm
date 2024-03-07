VERSION 5.00
Begin VB.Form frmVerSoporte 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "frmVerSoporte.frx":0000
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblR 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   3315
   End
   Begin VB.Image imgCerrar 
      Height          =   255
      Left            =   1800
      MouseIcon       =   "frmVerSoporte.frx":8EB7
      Top             =   3600
      Width           =   975
   End
End
Attribute VB_Name = "frmVerSoporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub

Private Sub imgCerrar_Click()
lblR.Caption = ""
Me.Hide
End Sub

