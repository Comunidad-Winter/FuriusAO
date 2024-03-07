VERSION 5.00
Begin VB.Form FrmGm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FrmGm"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtConsulta 
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar Consulta"
      Height          =   255
      Left            =   4440
      MouseIcon       =   "FrmGm.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione el tipo de consulta:"
      Height          =   3255
      Left            =   3480
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.OptionButton Opregunta 
         BackColor       =   &H000000FF&
         Caption         =   "Pregunta"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Obug 
         BackColor       =   &H000000FF&
         Caption         =   "Bug"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton OpersonajeT 
         BackColor       =   &H000000FF&
         Caption         =   "Personaje Trabado"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton Odescargas 
         BackColor       =   &H000000FF&
         Caption         =   "Descargas sobre baneos"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2760
         Width           =   2175
      End
      Begin VB.OptionButton Odenuncias 
         BackColor       =   &H000000FF&
         Caption         =   "Denuncias"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Texto de la consulta"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "FrmGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Private Sub cmdEnviar_Click()
Dim CONSULTA As String
CONSULTA = txtConsulta.Text
If txtConsulta.Text = "" Then
MsgBox "No ingresaste ningun texto."
Exit Sub
End If

If Odenuncias.value = True Then
Call SendData("µ" & CONSULTA)
ElseIf Opregunta.value = True Then
Call SendData("¶" & CONSULTA)
ElseIf Obug.value = True Then
Call SendData("¥" & CONSULTA)
ElseIf OpersonajeT.value = True Then
Call SendData("£" & CONSULTA)
ElseIf Odescargas.value = True Then
Call SendData("Å" & CONSULTA)
End If
Call SendData("/GS")
txtConsulta.Text = ""
Unload Me
End Sub

