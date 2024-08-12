VERSION 5.00
Begin VB.Form welcomeform 
   BackColor       =   &H00400000&
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10350
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exitbtn 
      Caption         =   "Salir "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9480
      TabIndex        =   3
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton regbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registrarse"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton logbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ingresar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label labelwelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "¡Bienvenido a Hub de libros!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
      Width           =   5055
   End
End
Attribute VB_Name = "welcomeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exitbtn_Click()
    End
End Sub

Private Sub Form_Load()
    'adjust the screen to the middle
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub logbtn_Click()
    loginform.Show
    welcomeform.Hide
    registerform.Hide
End Sub

Private Sub regbtn_Click()
    registerform.Show
    loginform.Hide
    welcomeform.Hide
End Sub
