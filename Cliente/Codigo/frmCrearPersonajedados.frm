VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   12075.47
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCorreo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      Left            =   720
      TabIndex        =   31
      Top             =   3000
      Width           =   2400
   End
   Begin VB.TextBox txtPasswdCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   33
      Top             =   765
      Width           =   2400
   End
   Begin VB.TextBox txtPasswd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   3831
      PasswordChar    =   "*"
      TabIndex        =   32
      Top             =   750
      Width           =   2400
   End
   Begin VB.TextBox txtCorreo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      Left            =   720
      TabIndex        =   30
      Top             =   1965
      Width           =   2400
   End
   Begin VB.ComboBox lstGenero 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":0000
      Left            =   720
      List            =   "frmCrearPersonajedados.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   5775
      Width           =   2400
   End
   Begin VB.ComboBox lstRaza 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":001D
      Left            =   720
      List            =   "frmCrearPersonajedados.frx":0030
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   4560
      Width           =   2400
   End
   Begin VB.ComboBox lstHogar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":005D
      Left            =   720
      List            =   "frmCrearPersonajedados.frx":0070
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   7140
      Width           =   2400
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   300
      Left            =   720
      MaxLength       =   20
      TabIndex        =   0
      Top             =   780
      Width           =   2415
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   1
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":00A1
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   38
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":01F3
      MousePointer    =   99  'Custom
      Top             =   7530
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   36
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":0345
      MousePointer    =   99  'Custom
      Top             =   7275
      Width           =   195
   End
   Begin VB.Label modCarisma 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5963
      TabIndex        =   47
      Top             =   7725
      Width           =   240
   End
   Begin VB.Label modInteligencia 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5963
      TabIndex        =   46
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label modConstitucion 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5963
      TabIndex        =   45
      Top             =   5910
      Width           =   240
   End
   Begin VB.Label modAgilidad 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5963
      TabIndex        =   44
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label modfuerza 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5963
      TabIndex        =   43
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   5040
      MouseIcon       =   "frmCrearPersonajedados.frx":0497
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblPass2OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   6360
      TabIndex        =   42
      Top             =   720
      Width           =   345
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   180
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblMailOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   3240
      TabIndex        =   38
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label lblMail2OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   3240
      TabIndex        =   36
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label lblPassOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   9960
      TabIndex        =   34
      Top             =   720
      Width           =   345
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   21
      Left            =   10785
      TabIndex        =   29
      Top             =   7950
      Width           =   405
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   42
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":1E19
      MousePointer    =   99  'Custom
      Top             =   8010
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   43
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":1F6B
      MousePointer    =   99  'Custom
      Top             =   8025
      Width           =   195
   End
   Begin VB.Label puntosquedan 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   10920
      TabIndex        =   27
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":20BD
      MousePointer    =   99  'Custom
      Top             =   3015
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   5
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":220F
      MousePointer    =   99  'Custom
      Top             =   3270
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   7
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":2361
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   9
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":24B3
      MousePointer    =   99  'Custom
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   11
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":2605
      MousePointer    =   99  'Custom
      Top             =   4020
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   13
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":2757
      MousePointer    =   99  'Custom
      Top             =   4260
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   15
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":28A9
      MousePointer    =   99  'Custom
      Top             =   4500
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   17
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":29FB
      MousePointer    =   99  'Custom
      Top             =   4770
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   19
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":2B4D
      MousePointer    =   99  'Custom
      Top             =   4995
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   21
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":2C9F
      MousePointer    =   99  'Custom
      Top             =   5265
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   23
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":2DF1
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   25
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":2F43
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   27
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":3095
      MousePointer    =   99  'Custom
      Top             =   6015
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   0
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":31E7
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   2
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":3339
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   4
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":348B
      MousePointer    =   99  'Custom
      Top             =   3255
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   6
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":35DD
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   8
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":372F
      MousePointer    =   99  'Custom
      Top             =   3765
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   10
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":3881
      MousePointer    =   99  'Custom
      Top             =   4020
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   12
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":39D3
      MousePointer    =   99  'Custom
      Top             =   4260
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   14
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":3B25
      MousePointer    =   99  'Custom
      Top             =   4500
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   16
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":3C77
      MousePointer    =   99  'Custom
      Top             =   4740
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   18
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":3DC9
      MousePointer    =   99  'Custom
      Top             =   4995
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   20
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":3F1B
      MousePointer    =   99  'Custom
      Top             =   5265
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   22
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":406D
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   24
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":41BF
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   26
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":4311
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   28
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":4463
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   29
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":45B5
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   30
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":4707
      MousePointer    =   99  'Custom
      Top             =   6525
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   31
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":4859
      MousePointer    =   99  'Custom
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   32
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":49AB
      MousePointer    =   99  'Custom
      Top             =   6750
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":4AFD
      MousePointer    =   99  'Custom
      Top             =   6750
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   34
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":4C4F
      MousePointer    =   99  'Custom
      Top             =   7020
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   35
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":4DA1
      MousePointer    =   99  'Custom
      Top             =   7050
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   37
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":4EF3
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   39
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":5045
      MousePointer    =   99  'Custom
      Top             =   7545
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   40
      Left            =   11160
      MouseIcon       =   "frmCrearPersonajedados.frx":5197
      MousePointer    =   99  'Custom
      Top             =   7785
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   10635
      MouseIcon       =   "frmCrearPersonajedados.frx":52E9
      MousePointer    =   99  'Custom
      Top             =   7815
      Width           =   195
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   1
      Left            =   240
      MouseIcon       =   "frmCrearPersonajedados.frx":543B
      MousePointer    =   99  'Custom
      Top             =   8400
      Width           =   1725
   End
   Begin VB.Image boton 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   0
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":6DBD
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   3120
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   20
      Left            =   10785
      TabIndex        =   26
      Top             =   7710
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   19
      Left            =   10785
      TabIndex        =   25
      Top             =   7470
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   18
      Left            =   10785
      TabIndex        =   24
      Top             =   7215
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   17
      Left            =   10785
      TabIndex        =   23
      Top             =   6960
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   16
      Left            =   10785
      TabIndex        =   22
      Top             =   6720
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   15
      Left            =   10785
      TabIndex        =   21
      Top             =   6465
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   14
      Left            =   10785
      TabIndex        =   20
      Top             =   6210
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   13
      Left            =   10785
      TabIndex        =   19
      Top             =   5955
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   12
      Left            =   10785
      TabIndex        =   18
      Top             =   5700
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   11
      Left            =   10785
      TabIndex        =   17
      Top             =   5445
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   10
      Left            =   10785
      TabIndex        =   16
      Top             =   5190
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   9
      Left            =   10785
      TabIndex        =   15
      Top             =   4935
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   8
      Left            =   10785
      TabIndex        =   14
      Top             =   4665
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   7
      Left            =   10785
      TabIndex        =   13
      Top             =   4425
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   6
      Left            =   10785
      TabIndex        =   12
      Top             =   4170
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   5
      Left            =   10785
      TabIndex        =   11
      Top             =   3930
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   4
      Left            =   10785
      TabIndex        =   10
      Top             =   3675
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   3
      Left            =   10785
      TabIndex        =   9
      Top             =   3420
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   2
      Left            =   10785
      TabIndex        =   8
      Top             =   3180
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   10785
      TabIndex        =   7
      Top             =   2700
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   10785
      TabIndex        =   6
      Top             =   2925
      Width           =   405
   End
   Begin VB.Label lbCarisma 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   435
      Left            =   5640
      TabIndex        =   5
      Top             =   7620
      Width           =   330
   End
   Begin VB.Label lbInteligencia 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   5640
      TabIndex        =   4
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label lbConstitucion 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   5640
      TabIndex        =   3
      Top             =   5805
      Width           =   495
   End
   Begin VB.Label lbAgilidad 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   5640
      TabIndex        =   2
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label lbFuerza 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   5640
      TabIndex        =   1
      Top             =   4080
      Width           =   495
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public SkillPoints As Byte
Function CheckData() As Boolean

If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserHogar = 0 Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If UserSexo = -1 Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True

End Function
Private Sub boton_Click(Index As Integer)
Dim i As Integer
Dim k As Object
        
Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0
        LlegoConfirmacion = False
        Confirmacion = 0

        i = 1
        
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
        UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
            UserName = Trim(UserName)
            MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.ListIndex + 1
        UserSexo = lstGenero.ListIndex
        UserHogar = lstHogar.ListIndex + 1
        
        UserAtributos(1) = 1
        UserAtributos(2) = 1
        UserAtributos(3) = 1
        UserAtributos(4) = 1
        UserAtributos(5) = 1
        
        If CheckData() Then
            UserPassword = MD5String(txtPasswd.Text)
            UserEmail = txtCorreo.Text
            
            If Not CheckMailString(UserEmail) Then
                MsgBox "Direccion de mail inválida.", vbExclamation, "Furius AO"
                txtCorreo.SetFocus
                Exit Sub
            End If
    
            If UserEmail <> txtCorreo2.Text Then
                MsgBox "Las direcciones de mail no coinciden.", vbExclamation, "Furius AO"
                txtCorreo2.Text = ""
                txtCorreo2.SetFocus
                Exit Sub
            End If
            
            If Len(Trim(txtPasswd)) = 0 Then
                MsgBox "Tenés que ingresar una contraseña.", vbExclamation, "Furius AO"
                txtPasswd.SetFocus
                Exit Sub
            End If
            
            If Len(Trim(txtPasswd)) < 6 Then
                MsgBox "El password debe tener al menos 6 caracteres.", vbExclamation, "Furius AO"
                txtPasswd = ""
                txtPasswdCheck = ""
                txtPasswd.SetFocus
                Exit Sub
            End If
            
            If Trim(txtPasswd) <> Trim(txtPasswdCheck) Then
                MsgBox "Las contraseñas no coinciden.", vbInformation, "Furius AO"
                txtPasswd = ""
                txtPasswdCheck = ""
                txtPasswd.SetFocus
                Exit Sub
            End If
    
            frmMain.Socket1.HostName = IPdelServidor
            frmMain.Socket1.RemotePort = PuertoDelServidor
    
            Me.MousePointer = 11
            EstadoLogin = CrearNuevoPj
    
            If Not frmMain.Socket1.Connected Then
                Call MsgBox("Error: Se ha perdido la conexion con el server.")
                Unload Me
            Else
                Call Login(ValidarLoginMSG(CInt(bRK)))
            End If
            
            If Musica = 0 Then
                CurMidi = DirMidi & "2.mid"
                LoopMidi = 1
                Call CargarMIDI(CurMidi)
                Call Play_Midi
            End If
        
            frmConnect.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        End If

    Case 1
        If Musica = 0 Then
            CurMidi = DirMidi & "2.mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If
        
        frmConnect.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        
        frmMain.Socket1.Disconnect
        frmConnect.MousePointer = 1
        Unload Me
End Select

End Sub
Private Sub Command1_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)

Dim indice
If Index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub
Private Sub Form_Load()

SkillPoints = 10
Puntos.Caption = SkillPoints
Me.Picture = LoadPicture(App.Path & "\graficos\Crearpersonaje.GIF")
Me.MousePointer = vbDefault

Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = ""
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = "+ 2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "- 3"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
        modCarisma.Caption = "- 3"
    Case Is = "Gnomo"
        modfuerza.Caption = "- 5"
        modAgilidad.Caption = "+ 4"
        modConstitucion.Caption = ""
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 1"
End Select

End Sub

Private Sub Pîcture4_Click()

End Sub

Private Sub Image1_Click()
PlayWaveDS (SND_CLICK)
Call SendData("TIRDAD")
End Sub

Private Sub lstRaza_click()

Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = ""
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = "+ 2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "- 3"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
        modCarisma.Caption = "- 3"
    Case Is = "Gnomo"
        modfuerza.Caption = "- 5"
        modAgilidad.Caption = "+ 4"
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 1"
End Select

End Sub

Private Sub txtCorreo_Change()

If Not CheckMailString(txtCorreo) Then
    lblMailOK = "O"
    lblMailOK.ForeColor = &HC0&
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
    Exit Sub
End If

lblMailOK = "P"
lblMailOK.ForeColor = &H80FF&

If (UCase$(txtCorreo.Text) = UCase$(txtCorreo2.Text)) Then
    lblMail2OK = "P"
    lblMail2OK.ForeColor = &H80FF&
Else
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
End If

End Sub
Private Sub txtCorreo_GotFocus()

MsgBox "La dirección de correo electrónico DEBE SER real."

End Sub
Private Sub txtCorreo2_Change()

If Not CheckMailString(txtCorreo) Then
    lblMailOK = "O"
    lblMailOK.ForeColor = &HC0&
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
    Exit Sub
End If

lblMailOK = "P"
lblMailOK.ForeColor = &H80FF&

If (UCase$(txtCorreo.Text) = UCase$(txtCorreo2.Text)) Then
    lblMail2OK = "P"
    lblMail2OK.ForeColor = &H80FF&
Else
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
End If

End Sub
Private Sub txtPasswd_Change()

If Len(Trim(txtPasswd)) < 6 Then
    lblPass2OK = "O"
    lblPass2OK.ForeColor = &HC0&
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
    Exit Sub
End If

lblPass2OK = "P"
lblPass2OK.ForeColor = &H80FF&

If (txtPasswdCheck = txtPasswd) Then
    lblPassOK = "P"
    lblPassOK.ForeColor = &H80FF&
Else
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
End If

End Sub
Private Sub txtPasswdCheck_Change()

If Len(Trim(txtPasswd)) < 6 Then
    lblPass2OK = "O"
    lblPass2OK.ForeColor = &HC0&
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
    Exit Sub
End If

lblPass2OK = "P"
lblPass2OK.ForeColor = &H80FF&

If (txtPasswdCheck = txtPasswd) Then
    lblPassOK = "P"
    lblPassOK.ForeColor = &H80FF&
Else
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
End If

End Sub
Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub
