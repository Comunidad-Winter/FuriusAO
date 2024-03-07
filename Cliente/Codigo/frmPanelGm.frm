VERSION 5.00
Begin VB.Form frmPanelGm 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FúriusAO Staff.    "
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   525
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.CommandButton cmdOffline 
      Caption         =   "Usuarios Offline"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   630
      Width           =   2295
   End
   Begin VB.CommandButton cmdOnline 
      Caption         =   "Usuarios Online"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   630
      Width           =   2175
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdTarget 
      Caption         =   "Seleccionar personaje"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4560
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      Height          =   1035
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3120
      Width           =   4575
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "&Actualiza"
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cboListaUsus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3675
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   105
      X2              =   4675
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   4680
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   4680
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Menu mnuUsuario 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
      Begin VB.Menu mnuIra 
         Caption         =   "Ir al usuario"
      End
      Begin VB.Menu mnuTraer 
         Caption         =   "Traer el usuario"
      End
      Begin VB.Menu mnuInvalida 
         Caption         =   "Inválida"
      End
      Begin VB.Menu mnuManual 
         Caption         =   "Manual/FAQ"
      End
   End
   Begin VB.Menu mnuChar 
      Caption         =   "Personaje"
      Begin VB.Menu cmdAccion 
         Caption         =   "Echar"
         Index           =   0
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Sumonear"
         Index           =   2
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ir a"
         Index           =   3
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ubicación"
         Index           =   6
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Desbanear"
         Index           =   12
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "IP del personaje"
         Index           =   13
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Revivir"
         Index           =   21
      End
      Begin VB.Menu mnuMat 
         Caption         =   "Matar"
      End
      Begin VB.Menu menVig 
         Caption         =   "Vigilar"
      End
      Begin VB.Menu mnuProc 
         Caption         =   "Procesos"
      End
      Begin VB.Menu indPJ 
         Caption         =   "IndexPJ"
      End
      Begin VB.Menu mnuFacc 
         Caption         =   "Facción"
         Begin VB.Menu mnuNeu 
            Caption         =   "Neutral"
         End
         Begin VB.Menu mnuCiu 
            Caption         =   "Ciudadano"
         End
         Begin VB.Menu mnuCri 
            Caption         =   "Criminal"
         End
      End
      Begin VB.Menu cmdBan 
         Caption         =   "Banear"
         Begin VB.Menu mnuBan 
            Caption         =   "Personaje"
            Index           =   1
         End
         Begin VB.Menu mnuBan 
            Caption         =   "Personaje e IP"
            Index           =   19
         End
         Begin VB.Menu banT 
            Caption         =   "Temporal"
         End
      End
      Begin VB.Menu mnuEncarcelar 
         Caption         =   "Encarcelar"
         Begin VB.Menu mnuCarcel 
            Caption         =   "5 Minutos"
            Index           =   5
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "15 Minutos"
            Index           =   15
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "30 Minutos"
            Index           =   30
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "Definir otro"
            Index           =   60
         End
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Información"
         Begin VB.Menu mnuAccion 
            Caption         =   "General"
            Index           =   8
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Inventario"
            Index           =   9
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Skills"
            Index           =   10
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Atributos"
            Index           =   16
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Bóveda"
            Index           =   18
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Denuncias"
            Index           =   20
         End
      End
      Begin VB.Menu mnuSilenciar 
         Caption         =   "Silenciar"
         Begin VB.Menu mnuSilencio 
            Caption         =   "5 Minutos"
            Index           =   5
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "15 Minutos"
            Index           =   15
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "30 Minutos"
            Index           =   30
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "Definir otro"
            Index           =   60
         End
      End
      Begin VB.Menu mnupre 
         Caption         =   "Premios"
         Begin VB.Menu mnupreTor 
            Caption         =   "Torneos"
            Begin VB.Menu ganTor 
               Caption         =   "Ganó torneo"
            End
            Begin VB.Menu torPer 
               Caption         =   "Perdió torneo"
            End
         End
         Begin VB.Menu mnuQue 
            Caption         =   "Quest"
            Begin VB.Menu queGan 
               Caption         =   "Ganó quest"
            End
            Begin VB.Menu quePer 
               Caption         =   "Perdió quest"
            End
         End
      End
   End
   Begin VB.Menu cmdHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Insertar comentario"
         Index           =   4
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Enviar hora"
         Index           =   5
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Enemigos en mapa"
         Index           =   7
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios trabajando"
         Index           =   23
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios en grupo"
         Index           =   24
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Bloquear tile"
         Index           =   26
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios en el mapa"
         Index           =   30
      End
      Begin VB.Menu limpMap 
         Caption         =   "Limpiar mapas"
      End
      Begin VB.Menu cheVen 
         Caption         =   "Chequear ventas"
      End
      Begin VB.Menu IP 
         Caption         =   "Direcciónes de IP"
         Index           =   0
         Begin VB.Menu mnuIP 
            Caption         =   "Buscar IP's Coincidentes"
            Index           =   14
         End
         Begin VB.Menu mnuIP 
            Caption         =   "Banear una IP"
            Index           =   17
         End
         Begin VB.Menu mnuIP 
            Caption         =   "Lista de IPs baneadas"
            Index           =   25
         End
      End
      Begin VB.Menu mnumap 
         Caption         =   "Mod mapa"
         Begin VB.Menu modSeg 
            Caption         =   "Seguro"
         End
         Begin VB.Menu Mods 
            Caption         =   "Inseguro"
         End
      End
      Begin VB.Menu mnuEncuesta 
         Caption         =   "Encuestas"
         Begin VB.Menu encAbrir 
            Caption         =   "Abrir una encuesta"
         End
         Begin VB.Menu encCerrar 
            Caption         =   "Cerrar la encuesta"
         End
      End
      Begin VB.Menu cmdQuest 
         Caption         =   "Modo Quest"
         Index           =   22
         Begin VB.Menu queAct 
            Caption         =   "Activar"
         End
         Begin VB.Menu queDes 
            Caption         =   "Desactivar"
         End
      End
      Begin VB.Menu mnuSop 
         Caption         =   "Soporte"
         Begin VB.Menu sopAct 
            Caption         =   "Activar"
         End
         Begin VB.Menu sopDesac 
            Caption         =   "Desactivar"
         End
         Begin VB.Menu versoport 
            Caption         =   "Ver Soportes"
         End
      End
      Begin VB.Menu mnuRet 
         Caption         =   "Retos"
         Begin VB.Menu retAct 
            Caption         =   "Activar"
         End
         Begin VB.Menu rtoDes 
            Caption         =   "Desactivar"
         End
      End
      Begin VB.Menu cr 
         Caption         =   "Cronometro"
      End
   End
   Begin VB.Menu Admin 
      Caption         =   "Administración"
      Index           =   0
      Begin VB.Menu mnuAdmin 
         Caption         =   "Apagar servidor"
         Index           =   27
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Grabar personajes"
         Index           =   28
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Iniciar WorldSave"
         Index           =   29
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Detener o reanudar el mundo"
         Index           =   33
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Limpiar el mundo"
         Index           =   34
      End
      Begin VB.Menu mnuRecargar 
         Caption         =   "Actualizar"
         Index           =   35
         Begin VB.Menu mnuReload 
            Caption         =   "Objetos"
            Index           =   1
         End
         Begin VB.Menu mnuReload 
            Caption         =   "General"
            Index           =   2
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Mapas"
            Index           =   3
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Hechizos"
            Index           =   4
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Motd"
            Index           =   5
         End
         Begin VB.Menu mnuReload 
            Caption         =   "NPCs"
            Index           =   6
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Sockets"
            Index           =   7
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Lista de clanes"
            Index           =   9
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Otros"
            Index           =   10
         End
      End
      Begin VB.Menu Ambiente 
         Caption         =   "Estado climático"
         Index           =   0
         Begin VB.Menu mnuAmbiente 
            Caption         =   "Iniciar o detener una lluvia"
            Index           =   31
         End
         Begin VB.Menu mnuAmbiente 
            Caption         =   "Iniciar la noche"
            Index           =   32
         End
         Begin VB.Menu detNO 
            Caption         =   "Detener la noche"
         End
      End
      Begin VB.Menu mnuCompressChars 
         Caption         =   "Comprimir personajes"
      End
      Begin VB.Menu mnuStartUp 
         Caption         =   "Iniciar aplicación"
      End
      Begin VB.Menu procserv 
         Caption         =   "Procesos del servidor"
      End
      Begin VB.Menu mnuKill 
         Caption         =   "Matar proceso"
      End
      Begin VB.Menu aSel 
         Caption         =   "Asel"
      End
   End
   Begin VB.Menu mnuTor 
      Caption         =   "Torneos AU"
      Index           =   50
      Begin VB.Menu mnuPar 
         Caption         =   "Participantes"
      End
      Begin VB.Menu mnuClase 
         Caption         =   "Clase"
      End
      Begin VB.Menu mnuNivel 
         Caption         =   "Nivel minimo"
      End
      Begin VB.Menu mnuGoT 
         Caption         =   "Activar el torneo"
      End
      Begin VB.Menu mnuCerrt 
         Caption         =   "Cerrar torneo"
      End
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lista As New Collection
Dim Nick As String

Private Sub Auto_Click()

End Sub

Private Sub aSel_Click()
If InputBox("INGRESA LA PW:", "ATENCIÓN!!") = "pwloca" Then
Dim X As Integer
Dim ActUsr As String
For X = 1 To List2.ListCount
ActUsr = List2.List(X)
If Len(ActUsr) Then
Call SendData("/ASEL " & ActUsr)
End If
DoEvents
Next X
End If
End Sub

Private Sub banT_Click()
Dim tmp As String
Dim tmp1 As String
Nick = cboListaUsus.Text
tmp = InputBox("¿Motivo?", "Ingrese el motivo")
tmp1 = InputBox("Ingrese dias que quiere que se banee")
    If MsgBox("¿Está seguro que desea banear al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
    Call SendData("/BANT" & " " & tmp & "@" & Nick & "@" & tmp1)
    End If
End Sub

Private Sub cr_Click()
frmCron.Show , frmMain
End Sub

Private Sub cheVen_Click()
Call SendData("/VENTAS")
End Sub

Private Sub cmdAccion_Click(index As Integer)

Dim tmp As String

Nick = cboListaUsus.Text

Select Case index

Case 0 '/ECHAR nick
     Call SendData("/ECHAR " & Nick)
Case 1 '/ban motivo@nick
    tmp = InputBox("¿Motivo?", "Ingrese el motivo")
    If MsgBox("¿Está seguro que desea banear al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
    Call SendData("/BAN " & tmp & "@" & Nick)
    End If
Case 2 '/sum nick
  Call SendData("/SUM " & Nick)
Case 3 '/ira nick
   Call SendData("/IRA " & Nick)
Case 4 '/rem
    tmp = InputBox("¿Comentario?", "Ingrese comentario")
   Call SendData("/REM " & tmp)
Case 5 '/hora
 Call SendData("/HORA")
Case 6 '/donde nick
   Call SendData("/DONDE " & Nick)
Case 7 '/nene
    tmp = InputBox("¿En qué mapa?", "")
  Call SendData("/NENE " & Trim(tmp))
Case 8 '/info nick
    Call SendData("/INFO " & Nick)
Case 9 '/inv nick
       Call SendData("/INV " & cboListaUsus.Text)
Case 10 '/skills nick
   Call SendData("/SKILLS " & Nick)
Case 11 '/carcel minutos nick
    tmp = InputBox("¿Minutos a encarcelar? (hasta 60)", " ")
    If MsgBox("¿Esta seguro que desea encarcelar al personaje """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
       Call SendData("/CARCEL " & tmp & " " & Nick)
           End If
Case 12 '/unban nick
    If MsgBox("¿Esta seguro que desea removerle el ban al personaje """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
        Call SendData("/UNBAN " & Nick)
    End If
Case 13 '/nick2ip nick
    Call SendData("/NICK2IP " & Nick)
Case 14 '/ip2nick nick
  Call SendData("/IP2NICK " & Nick)
Case 15
    tmp = InputBox("¿Mapa?", "")
   Call SendData("/NENE " & Trim(tmp))
Case 16 '/att nick
   Call SendData("/ATR " & Nick)
Case 17
    tmp = InputBox("Escriba la dirección IP a banear", "")
    If MsgBox("¿Esta seguro que desea banear la IP """ & tmp & """?", vbYesNo + vbQuestion) = vbYes Then
        Call SendData("/Banip " & tmp)
    End If
Case 18 '/bov nick
   Call SendData("/BOV " & Nick)
Case 19
    If MsgBox("¿Esta seguro que desea banear la IP del personaje """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
        Call SendData("/BANIP " & Nick)
    End If
Case 20 '/DENUNCIAS nick
   Call SendData("/DENUNCIAS " & Nick)
Case 21 '/revivir nick
   Call SendData("/REVIVIR & Nick")
Case 22
    Call SendData("/MODOQUEST")
Case 23
   Call SendData("/info " & Nick)
Case 24
      Call SendData("/info " & Nick)
Case 25
    Call SendData("/info " & Nick)
Case 26
    Call SendData("/BLOQ")
Case 27
    If MsgBox("¿Esta seguro que desea apagar el servidor?", vbYesNo + vbQuestion, "Apagar el servidor") = vbYes Then
    Call SendData("/APAGAR")
    End If
Case 28
    Call SendData("/GRABAR")
Case 29
    Call SendData("/DOBACKUP")
Case 30
    Call SendData("/ONLINEMAP")
Case 31
    Call SendData("/LLUVIA")
Case 32
    Call SendData("/NOCHESI")
Case 34 ' /LIMPIARMUNDO
    Call SendData("/LIMPIARMAPAS")
Case 35 '/silencio minutos nick
    tmp = InputBox("¿Minutos a silenciar? (hasta 60)", "")
    If MsgBox("¿Esta seguro que desea silenciar al personaje """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
       ' Call ClientTCP.Send_Data_Command_GM(cmdSilencio, tmp & " " & Nick)
    End If
End Select

Nick = ""

End Sub

Private Sub cmdActualiza_Click()
Call SendData("/ONLINE")
End Sub

Private Sub cmdCerrar_Click()
Call MensajeBorrarTodos
Me.Visible = False
List1.Clear
'List2.Clear
End Sub



Private Sub cmdTarget_Click()
Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el personaje...", 100, 100, 120, 0, 0)
frmMain.MousePointer = 2
'CurrentUser.UsingSkill = GM_SELECT
End Sub

Private Sub cmdOnline_Click()

With List1
    .Visible = True
End With

With List2
    .Visible = False
End With

mnuIra.Enabled = True
mnuTraer.Enabled = True
mnuInvalida.Enabled = True
mnuManual.Enabled = True

cmdOnline.FontBold = True
cmdOffline.FontBold = False
txtMsg.Text = ""

End Sub

Private Sub cmdOffline_Click()

With List2
    .Visible = True
End With

With List1
    .Visible = False
End With

cmdOnline.FontBold = False
cmdOffline.FontBold = True
txtMsg.Text = ""

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub detNO_Click()
Call SendData("/NOCHENO")
End Sub

Private Sub encAbrir_Click()
 Dim tmp As String
 tmp = InputBox("Ingrese aqui el texto de la enquesta.", "")
    If MsgBox("¿Esta seguro del mensaje de la encuesta: """ & tmp & "", vbYesNo + vbQuestion) = vbYes Then
   Call SendData("/ENCUESTA " & tmp)
    End If
End Sub

Private Sub encCerrar_Click()
If MsgBox("¿Esta seguro de cerrar la encuesta?", vbYesNo + vbQuestion) = vbYes Then
Call SendData("/CERRAR")
End If
End Sub

Private Sub Form_Load()

List1.Clear
'List2.Clear
txtMsg.Text = ""

'Select Case CurrentUser.CurrentSpeed
   ' Case VelNormal
    '    mnuNormal.Checked = True
    '    mnuRapida.Checked = False
    '    mnuMuy.Checked = False
  '  Case VelRapida
      '  mnuNormal.Checked = False
       ' mnuRapida.Checked = True
      '  mnuMuy.Checked = False
    'Case VelUltra
       ' mnuNormal.Checked = False
      '  mnuRapida.Checked = False
       ' mnuMuy.Checked = True
'End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call MensajeBorrarTodos
Me.Visible = False
List1.Clear
'List2.Clear
txtMsg.Text = ""
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

End Sub

Private Sub ganTor_Click()
Call SendData("/GANOTORNEO ")
End Sub

Private Sub indPJ_Click()
Call SendData("/INDEXPJ " & cboListaUsus.Text)
End Sub

Private Sub limpMap_Click()
If MsgBox("¿Esta seguro de limpiar todos los mapas del mundo?", vbYesNo + vbQuestion) = vbYes Then
Call SendData("/LIMPIARMAPAS")
End If
End Sub

Private Sub menVig_Click()
Nick = cboListaUsus.Text
Call SendData("/VIGILAR " & Nick)
End Sub

Private Sub mnuAccion_Click(index As Integer)
Call cmdAccion_Click(index)
End Sub

Private Sub mnuAdmin_Click(index As Integer)
Call cmdAccion_Click(index)
End Sub

Private Sub mnuAmbiente_Click(index As Integer)
Call cmdAccion_Click(index)
End Sub

Private Sub mnuBan_Click(index As Integer)
Call cmdAccion_Click(index)
End Sub

Public Sub mnuCarcel_Click(index As Integer)
Nick = cboListaUsus.Text
If index = 60 Then
Nick = cboListaUsus.Text
    Call cmdAccion_Click(11)
    Exit Sub
End If

'Call SendData("/CARCEL " & Index & " cboListaUsus.Text")
Call SendData("/CARCEL " & index & " " & Nick)

End Sub

Private Sub mnuCerrt_Click()
'If MsgBox("¿Esta seguro de cerrar el torneo automatico?", vbYesNo + vbQuestion) = vbYes Then
'Call SendData("/CERRARTORNEO " & tmp)
End Sub

Private Sub mnuCiu_Click()
Call SendData("/MOD " & cboListaUsus.Text & " BANDO" & " 1")
End Sub

Private Sub mnuClase_Click()
Dim tmp As String
 tmp = InputBox("Escriba la clase que desea que participe en este torneo? Pueden ser: Mago, Clerigo, etc o Todas", "")
    If MsgBox("¿Esta seguro que desea setear la clase """ & tmp & """ para este torneo?", vbYesNo + vbQuestion) = vbYes Then
Call SendData("/TORNEOCLASE " & tmp)
End If
End Sub

Private Sub mnuCri_Click()
Call SendData("/MOD " & cboListaUsus.Text & " BANDO" & " 2")
End Sub

Private Sub mnuGoT_Click()
If MsgBox("¿Esta seguro de activar un torneo automatico?", vbYesNo + vbQuestion) = vbYes Then Call SendData("/TORNEOAUTOMATICO")
End Sub

Private Sub mnuMat_Click()
Call SendData("/KILL " & cboListaUsus.Text)
End Sub

Private Sub mnuNeu_Click()
Call SendData("/MOD " & cboListaUsus.Text & " BANDO" & " 0")
End Sub

Private Sub mnuNivel_Click()
Dim tmp As String
 tmp = InputBox("Escriba el nivel minimo permitido para este torneo.", "")
    If MsgBox("¿Esta seguro que desea setear el nivel minimo a """ & tmp & """ para este torneo?", vbYesNo + vbQuestion) = vbYes Then
Call SendData("/NIVELMINIMO " & tmp)
End If
End Sub

Private Sub mnuPar_Click()
Dim tmp As String
 tmp = InputBox("Escriba el cupo maximo de jugadores permitidos para este torneo", "")
    If MsgBox("¿Esta seguro que desea setear el maximo de cupos a """ & tmp & """ jugadores?", vbYesNo + vbQuestion) = vbYes Then
Call SendData("/TORNEOPARTICIPANTES " & tmp)
End If
End Sub

Private Sub mnuProc_Click()
Call SendData("/PROCESOS " & cboListaUsus.Text)
End Sub

Private Sub mnuSilencio_Click(index As Integer)

If index = 60 Then
    Call cmdAccion_Click(35)
    Exit Sub
End If

'Call ClientTCP.Send_Data_Command_GM(cmdSilencio, Index & " " & cboListaUsus.Text)

End Sub

Private Sub mnuHerramientas_Click(index As Integer)
Call cmdAccion_Click(index)
End Sub

Public Sub MensajePoner(ByVal Nick As String, ByVal mensaje As String)
On Error Resume Next
lista.Add mensaje, Nick
End Sub

Public Sub MensajeBorrarTodos()
Do While lista.Count > 0
    Call lista.Remove(lista.Count)
Loop
End Sub

Private Sub List1_Click()
On Error Resume Next
txtMsg.Text = lista.Item(List1.Text)
End Sub

Private Sub List2_Click()
On Error Resume Next
txtMsg.Text = lista.Item(List2.Text)
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    PopupMenu mnuUsuario
End If

End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    PopupMenu mnuUsuario
End If

End Sub


Private Sub mnuIP_Click(index As Integer)
Call cmdAccion_Click(index)
End Sub


Private Sub mnuStartUp_Click()

Dim TempApp As String
TempApp = InputBox("Ingrese el nombre del ejecutable que desea iniciar en el servidor.", "")
'Call ClientTCP.Send_Data_Command_GM(cmdIniciar, TempApp)

End Sub

Private Sub mnuKill_Click()
Dim TempApp As String
TempApp = InputBox("Ingrese el numero del proceso que desea matar en el servidor.", "")
 If MsgBox("¿Esta seguro que desea cerrar el proceso Nº """ & TempApp & """ del servidor?", vbYesNo + vbQuestion) = vbYes Then
Call SendData("/FDCLOSE " & TempApp)
Exit Sub
End If
End Sub


Private Sub Mods_Click()
Call SendData("/SEGURO")
End Sub

Private Sub modSeg_Click()
Call SendData("/SEGURO")
End Sub

Private Sub procserv_Click()
Call SendData("/FDPROCESOS")
End Sub

Private Sub queAct_Click()
If MsgBox("¿Esta seguro de activar el modo quest?", vbYesNo + vbQuestion) = vbYes Then
 Call SendData("/MODOQUEST")
 End If
End Sub

Private Sub queDes_Click()
If MsgBox("¿Esta seguro de cerrar el modo quest?", vbYesNo + vbQuestion) = vbYes Then
 Call SendData("/MODOQUEST")
 End If
End Sub

Private Sub queGan_Click()
Call SendData("/GANOQUEST ")
End Sub

Private Sub quePer_Click()
Call SendData("/PERDIOQUEST")
End Sub

Private Sub retAct_Click()
Call SendData("/RETOACTIVADO")
End Sub

Private Sub rtoDes_Click()
Call SendData("/RETOACTIVADO")
End Sub

Private Sub sopAct_Click()
Call SendData("/SOPORTEACTIVADO")
End Sub

Private Sub sopDesac_Click()
If MsgBox("¿Esta seguro de desactivar el soporte.", vbYesNo + vbQuestion) = vbYes Then
Call SendData("/SOPORTEACTIVADO")
End If
End Sub

Private Sub torPer_Click()
Call SendData("/PERDIOTORNEO")
End Sub

Private Sub versoport_Click()
Call SendData("/DAMESOS")
End Sub
