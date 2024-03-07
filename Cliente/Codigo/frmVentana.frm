VERSION 5.00
Begin VB.Form frmVentana 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ventana"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   3375
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Aplicaciones 
      AutoSize        =   -1  'True
      Caption         =   "Aplicaciones"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Procesos 
      AutoSize        =   -1  'True
      Caption         =   "Procesos"
      Height          =   195
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmVentana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_LostFocus()

Me.Visible = False

End Sub

