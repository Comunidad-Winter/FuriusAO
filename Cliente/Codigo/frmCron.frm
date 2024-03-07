VERSION 5.00
Begin VB.Form frmCron 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Cronometro GM"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   1980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   1980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotonCmd 
      Caption         =   "Mostrar Tiempo"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   0
   End
   Begin VB.CommandButton BotonCmd 
      Caption         =   "PAUSE"
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton BotonCmd 
      Caption         =   "PLAY"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton BotonCmd 
      Caption         =   "STOP"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton BotonCmd 
      Caption         =   "YA"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblTiempo 
      Alignment       =   2  'Center
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "frmCron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SegundosP As Integer
Private Sub BotonCmd_Click(index As Integer)

Select Case index

    Case 0
        tmrTime.Enabled = True
        SegundosP = 0
    Case 1 'stop
        SegundosP = 0
        lblTiempo.Caption = Format(Int(SegundosP / 3600) Mod 24, "00") & ":" & _
                    Format(Int(SegundosP / 60) Mod 60, "00") & ":" & _
                    Format(Int(SegundosP) Mod 60, "00")
        
        tmrTime.Enabled = False
    Case 2 'play
        tmrTime.Enabled = True
    Case 3 'pause
        tmrTime.Enabled = False
    Case 4 'show
        Call SendData("/RMSG TIEMPO:" & lblTiempo.Caption & "~255~255~255~1~0")
End Select

End Sub

Private Sub tmrTime_Timer()
SegundosP = SegundosP + 1
lblTiempo.Caption = Format(Int(SegundosP / 3600) Mod 24, "00") & ":" & _
                    Format(Int(SegundosP / 60) Mod 60, "00") & ":" & _
                    Format(Int(SegundosP) Mod 60, "00")
'1000 EN UN SEGUNDO

End Sub
