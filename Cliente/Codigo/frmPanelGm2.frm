VERSION 5.00
Begin VB.Form frmPanelGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel GM"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5595
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "Dungeon Furius"
      Height          =   375
      Left            =   2040
      TabIndex        =   30
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DungeonVeril"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MagiaHilidan"
      Height          =   375
      Left            =   3720
      TabIndex        =   28
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MagiaLindos"
      Height          =   375
      Left            =   2040
      TabIndex        =   27
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MagiaBander"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton torneo1 
      Caption         =   "VER JUGADORES"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   25
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton torneo 
      Caption         =   "A/C TORNEO"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   24
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdAcction 
      Caption         =   "Gm's Online"
      Height          =   315
      Index           =   19
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Show SOS"
      Height          =   315
      Index           =   18
      Left            =   3420
      TabIndex        =   21
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Boveda"
      Height          =   315
      Index           =   17
      Left            =   2340
      TabIndex        =   20
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Ban X ip"
      Height          =   315
      Index           =   16
      Left            =   1260
      TabIndex        =   19
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Denuncias"
      Height          =   315
      Index           =   15
      Left            =   180
      TabIndex        =   18
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "IP 2 NICK"
      Height          =   315
      Index           =   14
      Left            =   1260
      TabIndex        =   17
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "NICK 2 IP"
      Height          =   315
      Index           =   13
      Left            =   180
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "UNBAN"
      Height          =   315
      Index           =   12
      Left            =   3420
      TabIndex        =   15
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "CARCEL"
      Height          =   315
      Index           =   11
      Left            =   3420
      TabIndex        =   14
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "SKILLS"
      Height          =   315
      Index           =   10
      Left            =   1260
      TabIndex        =   13
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "INV"
      Height          =   315
      Index           =   9
      Left            =   180
      TabIndex        =   12
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "INFO"
      Height          =   315
      Index           =   8
      Left            =   3420
      TabIndex        =   11
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "N.ENE."
      Height          =   315
      Index           =   7
      Left            =   180
      TabIndex        =   10
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "DONDE"
      Height          =   315
      Index           =   6
      Left            =   3420
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "HORA"
      Height          =   315
      Index           =   5
      Left            =   2340
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Guardar comentario"
      Height          =   315
      Index           =   4
      Left            =   180
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "IRA"
      Height          =   315
      Index           =   3
      Left            =   1260
      TabIndex        =   6
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "SUM"
      Height          =   315
      Index           =   2
      Left            =   2340
      TabIndex        =   5
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "BAN"
      Height          =   315
      Index           =   1
      Left            =   2340
      TabIndex        =   4
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "ECHAR"
      Height          =   315
      Index           =   0
      Left            =   2340
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "Actualiza"
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   4560
      Width           =   4035
   End
   Begin VB.ComboBox cboListaUsus 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3435
   End
   Begin VB.Label Label1 
      Caption         =   "FuriusAO"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2880
      TabIndex        =   23
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   120
      X2              =   120
      Y1              =   540
      Y2              =   1380
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   4440
      X2              =   4440
      Y1              =   540
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2280
      X2              =   2280
      Y1              =   960
      Y2              =   1380
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2280
      X2              =   4440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4440
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   2280
      Y1              =   1380
      Y2              =   1380
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.2
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Private Sub cmdAccion_Click(Index As Integer)
Dim Ok As Boolean, Tmp As String, Tmp2 As String
Dim Nick As String

Nick = cboListaUsus.Text

Select Case Index
Case 0 '/ECHAR nick
    Call SendData("/ECHAR " & Nick)
Case 1 '/ban motivo@nick
    Tmp = InputBox("Motivo ?", "")
    If MsgBox("Esta seguro que desea banear al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/BAN " & Tmp & "@" & Nick)
    End If
Case 2 '/sum nick
    Call SendData("/SUM " & Nick)
Case 3 '/ira nick
    Call SendData("/IRA " & Nick)
Case 4 '/rem
    Tmp = InputBox("Comentario ?", "")
    Call SendData("/REM " & Tmp)
Case 5 '/hora
    Call SendData("/HORA")
Case 6 '/donde nick
    Call SendData("/DONDE " & Nick)
Case 7 '/nene
    Tmp = InputBox("Mapa ?", "")
    Call SendData("/NENE " & Trim(Tmp))
Case 8 '/info nick
    Call SendData("/INFO " & Nick)
Case 9 '/inv nick
    Call SendData("/INV " & Nick)
Case 10 '/skills nick
    Call SendData("/SKILLS " & Nick)
Case 11 '/carcel minutos nick
    Tmp = InputBox("Minutos ? (hasta 30)", "")
    Tmp2 = InputBox("Razon ?", "")
    If MsgBox("Esta seguro que desea encarcelar al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/CARCEL " & Nick & "@" & Tmp2 & "@" & Tmp)
    End If
Case 12 '/unban nick
    If MsgBox("Esta seguro que desea removerle el ban al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/UNBAN " & Nick)
    End If
Case 13 '/nick2ip nick
    Call SendData("/NICKIP " & Nick)
Case 14 '/ip2nick ip
    Call SendData("/IPNICK " & Nick)
Case 15 '/Denuncias
    Call SendData("/DENUNCIAS " & cboListaUsus.Text)
Case 16 'Ban X ip
    If MsgBox("Esta seguro que desea banear el (ip o personaje) " & Nick & "Por IP?", vbYesNo) = vbYes Then
    Call SendData("/BANIP " & Nick)
    End If
Case 17 ' MUESTA BOBEDA
    Call SendData("/BOV " & Nick)
Case 18 ' Sos
    Call SendData("/SHOW SOS")
Case 19 ' GMS ONline
    Call SendData("/ONLINEGM")
End Select
End Sub

Private Sub cmdAcction_Click(Index As Integer)
    Call SendData("/ONLINEGM")
End Sub

Private Sub cmdActualiza_Click()
Call SendData("LISTUSU")

End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Call SendData("/CT 59 90 63")
Call SendData("/RACC 219")
End Sub

Private Sub Command2_Click()
Call SendData("/CT 62 50 77")
Call SendData("/RACC 220")
End Sub

Private Sub Command3_Click()
Call SendData("/CT 149 28 68")
Call SendData("/RACC 218")
End Sub

Private Sub Command4_Click()
Call SendData("/CT 139 50 50")
Call SendData("/RACC 217")
End Sub

Private Sub Command5_Click()
Call SendData("/CT 169 45 45")
Call SendData("/RACC 216")
End Sub

Private Sub Form_Load()
Me.Show
Call cmdActualiza_Click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Skills_Click(Index As Integer)
    Tmp = InputBox("Skills ? (hasta 100)", "cboListaUsus.Text")
        If MsgBox("Esta seguro que desea editar los skills al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/MOD " & Nick & "SKI" & Tmp)
End If
End Sub

Private Sub torneo_Click(Index As Integer)
Call SendData("/TORNEO")
End Sub

Private Sub torneo1_Click(Index As Integer)
Call SendData("/VERTORNEO")
End Sub
