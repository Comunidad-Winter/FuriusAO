VERSION 5.00
Begin VB.Form frmCustomKeys 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Teclados"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Teclado 6"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   6
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Teclado 5"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   5
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Teclado 4"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   4
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Teclado 3"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   3
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Teclado 2"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Teclado 1"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   1
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Defecto"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar y Salir"
      Height          =   375
      Left            =   3480
      TabIndex        =   67
      Top             =   6360
      Width           =   4575
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Otros"
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   3480
      TabIndex        =   52
      Top             =   3240
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   27
         Left            =   1920
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   1920
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   26
         Left            =   1920
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   1920
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   1920
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   1920
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   1920
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Capturar Pantalla"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Macro Trabajo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Macro Hechizos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Meditar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar Opciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar/Ocultar FPS"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Guild - Clan"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3480
      TabIndex        =   24
      Top             =   2160
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   1560
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   1560
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedir Ayuda"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seguro de Clan"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Acciones"
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      TabIndex        =   29
      Top             =   2520
      Width           =   3255
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   1440
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   1440
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   1440
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   1440
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   1440
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   1440
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1440
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1440
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1440
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seguro Resu"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Atacar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tirar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ocultar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modo Seguro"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Robar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Domar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modo Combate"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Agarrar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Opciones Personales"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3480
      TabIndex        =   17
      Top             =   720
      Width           =   4335
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2520
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   2520
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2520
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar/Ocultar Nombres"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Corregir Posicion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Activar/Desactivar Musica"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Movimiento"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   3255
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Derecha"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Izquierda"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Abajo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Arriba"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

''
'frmCustomKeys - Allows the user to customize keys.
'Implements class clsCustomKeys
'
'@author Rapsodius
'@date 20070805
'@version 1.2.0
'@see clsCustomKeys

Option Explicit
Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim i As Long
    
    Check1(Index).value = vbChecked
    CustomKeys.KeybConfig = Index

    For i = 1 To CustomKeys.count
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
        If Index = 0 Then
            Text1(i).Enabled = False
        Else
            Text1(i).Enabled = True
        End If
    Next i
        
    For i = 0 To CustomKeys.KeyCount
        If i <> Index Then
            Check1(i).value = vbUnchecked
        End If
    Next i
End Sub

Private Sub Command1_Click()

Dim i As Long

    'Call Audio.PlayWave(SND_CLICK)

    For i = 1 To CustomKeys.count
        If LenB(Text1(i).Text) = 0 Then
            Call MsgBox("Hay una o mas teclas no validas, por favor verifique.", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Argentum Online")
            Exit Sub
        End If
    Next i
   'Call SaveCustomKeys
    Unload Me
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim i As Long

    Check1(CustomKeys.KeybConfig).value = vbChecked

    For i = 1 To CustomKeys.count
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
        If CustomKeys.KeybConfig = 0 Then
            Text1(i).Enabled = False
        Else
            Text1(i).Enabled = True
        End If
    Next i
    
    For i = 0 To CustomKeys.KeyCount
        If i <> CustomKeys.KeybConfig Then
            Check1(i).value = vbUnchecked
        End If
    Next i
    
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Long

    If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
    'If key is not valid, we exit
    
    Text1(Index).Text = CustomKeys.ReadableName(KeyCode)
    Text1(Index).SelStart = Len(Text1(Index).Text)
    
    For i = 1 To CustomKeys.count
        If i <> Index Then
            If CustomKeys.BindedKey(i) = KeyCode Then
                Text1(Index).Text = "" 'If the key is already assigned, simply reject it
                Call Beep 'Alert the user
                KeyCode = 0
                Exit Sub
            End If
        End If
    Next i
    
    CustomKeys.BindedKey(Index) = KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call Text1_KeyDown(Index, KeyCode, Shift)
End Sub
