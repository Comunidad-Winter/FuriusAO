VERSION 5.00
Begin VB.Form frmHlp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ayuda"
   ClientHeight    =   4200
   ClientLeft      =   2355
   ClientTop       =   1845
   ClientWidth     =   5655
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2898.914
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   5310.338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFAQ 
      Caption         =   "Preguntas &frecuentes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdBPage 
      Caption         =   "< Página &anterior"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdNPage 
      Caption         =   "Página &siguiente >"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame Ayuda 
      Caption         =   "Introducción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.Label lblHlp 
         Height          =   2715
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   5085
      End
   End
End
Attribute VB_Name = "frmHlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EnQuePagina As Integer

Private Sub cmdFAQ_Click()
    lblHlp.Caption = "."
End Sub

Private Sub cmdBPage_Click()
    EnQuePagina = EnQuePagina - 1
    Call CambioPagina
End Sub

Private Sub cmdNPage_Click()
    EnQuePagina = EnQuePagina + 1
    Call CambioPagina
End Sub

Private Sub CambioPagina()

Select Case EnQuePagina

Case 0
    Ayuda.Caption = "Introducción"
    lblHlp.Caption = "En FúriusAo encontrarás todo un nuevo mundo por explorar, sin fronteras, en el que no hay profesión que sobresalga sobre otra y aún el ambiente se mantiene vivo. A continuación podés encontrar algunas de las preguntas más frecuentes que se realizan los viajeros de estas tierras."
    cmdBPage.Visible = False
    cmdNPage.Visible = True
Case 1
    Ayuda.Caption = "Empezando a jugar"
    lblHlp.Caption = "En un principio, tendrás, como mínimo, ropa, agua, comida y un arma. Un buen lugar para empezar es el Newbie Dungeon, que posee dos zonas. En la primera encontrarás tenebrosas criaturas que te desafiarán, en la segunda una zona en donde entrenar tus habilidades de combate. Estando un alguna de las ciudades iniciales, podrás tomar el teleport al mismo fácilmente. Mientras subas de nivel, irás ganando oro, con lo que puedes comprar hechizos, armas, ropa, etc. Cuando tu nivel sea 13 o superior, dejarás de ser newbie y ya no ganarás oro al subir de nivel. Deberás ir pensando en acceder a los bosques, los dungeons, y todo el mundo que te espera."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 2
    Ayuda.Caption = "Entrenando"
    lblHlp.Caption = "Cada clase es única. Al haber elegido cierta clase, estas sujeto a ser mejor en algunas áreas que en otras. Aprovéchalas. Si eres mago, dedícate a la magia. Si eres guerrero, dedícate al combate con armas. Cuando subas de nivel, ganarás skillpoints, los cuales puedes decidir utilizar o no. Lo más conveniente es dejarlos para cuando seas de mayor nivel. Por ahora, mientras entrenes, irás ganando naturalmente los skillpoints. A mayor skill en tal área, mejor es el rendimiento que tendrá tu personaje al realizar la actividad. Explora: es la mejor manera de conocer el mundo en el que vive tu personaje. Encontrarás miles de criaturas, nuevos retos y gente con quien afrontarlos. Los mejores lugares para comenzar a entrenar luego de haber dejado el Newbie Dungeon son el Bosque Dorck (Mapas 39 y 38), En estas zonas encontrarás tus primeros retos: Arañas Gigantes y Zombies."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 3
    Ayuda.Caption = "Interactuando"
    lblHlp.Caption = "A lo largo de tus viajes, encontrarás muchas maneras de interactuar con el mundo en el que tu personaje vive. Encontrarás ciudades, dungeons, bosques, y muchos lugares que llamarán tu atención. Si estás en una ciudad una buena idea es, si es que ésta tiene un puerto o río cercano, es pescar. Puedes adquirir una caña o red de pesca en el gremio de pescadores. Otra opción es talar. A medida que explores las diferentes maneras de ganarte la vida, decidirás cuál crees como más apropiada. No dejes de probar la minería y herrería, pues es una actividad muy productiva e útil si es que tu personaje tiene habilidades de combate. En FúriusAo estas áreas han sido rebalanceadas para obtener, así, un mundo más justo y próspero."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 4
    Ayuda.Caption = "Facciónes"
    lblHlp.Caption = "En la medida de que llegues a niveles superiores, podrás decidir enlistarte o no en una facción. Puedes elegir las Tropas Reales o bien las Legiones del Caos. Para ingresar en las Tropas Reales debes haber matado, como mínimo, cincuenta criminales. Si en algún momento asesinaste gente inocente no serás admitido. Tu nivel no deberá ser inferior a 20. Por otro lado, si tu deseo es ingresar en las Legiones del Caos, deberás haber matado como mínimo 50 ciudadanos y ser nivel 25 o superior. Cuando llegues a la centena de criminales o ciudadanos asesinados, serás recompensado y, además de ganar algo de experiencia, subirás un rango de jerarquía, lo que, además de posibilitarte nuevos ítems especiales, hará que tu rango sea superior. Recuerda, no hay vuelta atrás a la hora de seleccionar una facción. Se cuidadoso a la hora de elegir."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 5
    Ayuda.Caption = "Magia"
    lblHlp.Caption = "FúriusAo fue creado ni mas ni menos que con la quinta esencia, lo que hace a la magia y a lo desconocido. Aprovéchala. Está en todas partes, la materialización, y mucho más que tienes por explorar. Se dice que hay magos con poderes tales, que pueden abrir portales luminosos que con su magia para viajar entre el espacio. Muchos otros dicen haber visto magia con poder suficiente para matar hasta al más poderoso guerrero de estas tierras."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 6
    Ayuda.Caption = "Viajando"
    lblHlp.Caption = "Hay varias maneras no convencionales de viajar hacia otras tierras. Si hay un puerto en las cercanías y quieres llegar rápido a tu destino ¿Por qué no tomar un barco? Es muy fácil, sólo debes acercarte el muelle, comerciar con el pirata y luego mostrarle el pasaje, así, llegarás al destino elegido. Los pases caen al morir ser robados. Lleva un mapa. Te salvará un situaciónes comprometedoras. Recuerda que para viajar a zonas lejanas debes poder usar una Barca."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 7
    Ayuda.Caption = "Objetos mágicos"
    lblHlp.Caption = "Anillos, Baculo Sagrado y todo tipo de extraños objetos han sido impregnados de magia por los más poderosos y Arcanos hechizeros."
    cmdBPage.Visible = True
    cmdNPage.Visible = True

Case 8
    Ayuda.Caption = "Game Masters"
    lblHlp.Caption = "Este mundo se encuentra regido por dioses, quienes mantienen el órden y están constantemente trabajando por su mejoría. Puedes invocarlos escribiendo /soporte. De este modo, los dioses acudirán en tu ayuda. Los mismos son fácilmente reconocibles, su nombre, habla, y descripción se encuentra en un color no muy comun."
    cmdBPage.Visible = True
    cmdNPage.Visible = False
End Select
End Sub

