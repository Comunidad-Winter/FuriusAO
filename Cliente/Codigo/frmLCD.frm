VERSION 5.00
Begin VB.Form frmLCD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "   mcLCD Sample App"
   ClientHeight    =   510
   ClientLeft      =   4545
   ClientTop       =   3675
   ClientWidth     =   2040
   ControlBox      =   0   'False
   Icon            =   "frmLCD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   510
   ScaleWidth      =   2040
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   2115
      Top             =   45
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF8080&
      Height          =   510
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   1980
      TabIndex        =   0
      Top             =   0
      Width           =   2040
   End
End
Attribute VB_Name = "frmLCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lcdTest As New mcLCD
Private Sub Form_Load()

   lcdTest.NewLCD Picture1

End Sub


Private Sub Picture1_Click()

   Unload frmLCD

End Sub

Private Sub Timer1_Timer()

   lcdTest.Caption = Time

End Sub


