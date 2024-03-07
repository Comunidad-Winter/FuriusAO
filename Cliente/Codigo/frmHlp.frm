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
      Caption         =   "< P�gina &anterior"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdNPage 
      Caption         =   "P�gina &siguiente >"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame Ayuda 
      Caption         =   "Introducci�n"
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
    Ayuda.Caption = "Introducci�n"
    lblHlp.Caption = "En F�riusAo encontrar�s todo un nuevo mundo por explorar, sin fronteras, en el que no hay profesi�n que sobresalga sobre otra y a�n el ambiente se mantiene vivo. A continuaci�n pod�s encontrar algunas de las preguntas m�s frecuentes que se realizan los viajeros de estas tierras."
    cmdBPage.Visible = False
    cmdNPage.Visible = True
Case 1
    Ayuda.Caption = "Empezando a jugar"
    lblHlp.Caption = "En un principio, tendr�s, como m�nimo, ropa, agua, comida y un arma. Un buen lugar para empezar es el Newbie Dungeon, que posee dos zonas. En la primera encontrar�s tenebrosas criaturas que te desafiar�n, en la segunda una zona en donde entrenar tus habilidades de combate. Estando un alguna de las ciudades iniciales, podr�s tomar el teleport al mismo f�cilmente. Mientras subas de nivel, ir�s ganando oro, con lo que puedes comprar hechizos, armas, ropa, etc. Cuando tu nivel sea 13 o superior, dejar�s de ser newbie y ya no ganar�s oro al subir de nivel. Deber�s ir pensando en acceder a los bosques, los dungeons, y todo el mundo que te espera."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 2
    Ayuda.Caption = "Entrenando"
    lblHlp.Caption = "Cada clase es �nica. Al haber elegido cierta clase, estas sujeto a ser mejor en algunas �reas que en otras. Aprov�chalas. Si eres mago, ded�cate a la magia. Si eres guerrero, ded�cate al combate con armas. Cuando subas de nivel, ganar�s skillpoints, los cuales puedes decidir utilizar o no. Lo m�s conveniente es dejarlos para cuando seas de mayor nivel. Por ahora, mientras entrenes, ir�s ganando naturalmente los skillpoints. A mayor skill en tal �rea, mejor es el rendimiento que tendr� tu personaje al realizar la actividad. Explora: es la mejor manera de conocer el mundo en el que vive tu personaje. Encontrar�s miles de criaturas, nuevos retos y gente con quien afrontarlos. Los mejores lugares para comenzar a entrenar luego de haber dejado el Newbie Dungeon son el Bosque Dorck (Mapas 39 y 38), En estas zonas encontrar�s tus primeros retos: Ara�as Gigantes y Zombies."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 3
    Ayuda.Caption = "Interactuando"
    lblHlp.Caption = "A lo largo de tus viajes, encontrar�s muchas maneras de interactuar con el mundo en el que tu personaje vive. Encontrar�s ciudades, dungeons, bosques, y muchos lugares que llamar�n tu atenci�n. Si est�s en una ciudad una buena idea es, si es que �sta tiene un puerto o r�o cercano, es pescar. Puedes adquirir una ca�a o red de pesca en el gremio de pescadores. Otra opci�n es talar. A medida que explores las diferentes maneras de ganarte la vida, decidir�s cu�l crees como m�s apropiada. No dejes de probar la miner�a y herrer�a, pues es una actividad muy productiva e �til si es que tu personaje tiene habilidades de combate. En F�riusAo estas �reas han sido rebalanceadas para obtener, as�, un mundo m�s justo y pr�spero."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 4
    Ayuda.Caption = "Facci�nes"
    lblHlp.Caption = "En la medida de que llegues a niveles superiores, podr�s decidir enlistarte o no en una facci�n. Puedes elegir las Tropas Reales o bien las Legiones del Caos. Para ingresar en las Tropas Reales debes haber matado, como m�nimo, cincuenta criminales. Si en alg�n momento asesinaste gente inocente no ser�s admitido. Tu nivel no deber� ser inferior a 20. Por otro lado, si tu deseo es ingresar en las Legiones del Caos, deber�s haber matado como m�nimo 50 ciudadanos y ser nivel 25 o superior. Cuando llegues a la centena de criminales o ciudadanos asesinados, ser�s recompensado y, adem�s de ganar algo de experiencia, subir�s un rango de jerarqu�a, lo que, adem�s de posibilitarte nuevos �tems especiales, har� que tu rango sea superior. Recuerda, no hay vuelta atr�s a la hora de seleccionar una facci�n. Se cuidadoso a la hora de elegir."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 5
    Ayuda.Caption = "Magia"
    lblHlp.Caption = "F�riusAo fue creado ni mas ni menos que con la quinta esencia, lo que hace a la magia y a lo desconocido. Aprov�chala. Est� en todas partes, la materializaci�n, y mucho m�s que tienes por explorar. Se dice que hay magos con poderes tales, que pueden abrir portales luminosos que con su magia para viajar entre el espacio. Muchos otros dicen haber visto magia con poder suficiente para matar hasta al m�s poderoso guerrero de estas tierras."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 6
    Ayuda.Caption = "Viajando"
    lblHlp.Caption = "Hay varias maneras no convencionales de viajar hacia otras tierras. Si hay un puerto en las cercan�as y quieres llegar r�pido a tu destino �Por qu� no tomar un barco? Es muy f�cil, s�lo debes acercarte el muelle, comerciar con el pirata y luego mostrarle el pasaje, as�, llegar�s al destino elegido. Los pases caen al morir ser robados. Lleva un mapa. Te salvar� un situaci�nes comprometedoras. Recuerda que para viajar a zonas lejanas debes poder usar una Barca."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 7
    Ayuda.Caption = "Objetos m�gicos"
    lblHlp.Caption = "Anillos, Baculo Sagrado y todo tipo de extra�os objetos han sido impregnados de magia por los m�s poderosos y Arcanos hechizeros."
    cmdBPage.Visible = True
    cmdNPage.Visible = True

Case 8
    Ayuda.Caption = "Game Masters"
    lblHlp.Caption = "Este mundo se encuentra regido por dioses, quienes mantienen el �rden y est�n constantemente trabajando por su mejor�a. Puedes invocarlos escribiendo /soporte. De este modo, los dioses acudir�n en tu ayuda. Los mismos son f�cilmente reconocibles, su nombre, habla, y descripci�n se encuentra en un color no muy comun."
    cmdBPage.Visible = True
    cmdNPage.Visible = False
End Select
End Sub

