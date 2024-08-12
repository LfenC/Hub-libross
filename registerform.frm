VERSION 5.00
Begin VB.Form registerform 
   BackColor       =   &H00400000&
   Caption         =   "Register form"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9060
   FillColor       =   &H00400000&
   ForeColor       =   &H00400000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton backtbn 
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton resetbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Restablecer"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   10
      Top             =   4560
      Width           =   2055
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
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   9
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox passwordtext 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox usernametext 
      Height          =   405
      Left            =   3120
      TabIndex        =   6
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox lastnametext 
      Height          =   405
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox nametext 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label passwordlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label usernamelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de usuario:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lastnamelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label namelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label registerlabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registrarse"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "registerform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub backtbn_Click()
    welcomeform.Show
    Unload Me
End Sub

Private Sub Form_Load()
    If con Is Nothing Then
        Connect
    ElseIf con.State = adStateClosed Then
        Connect
    End If
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    rs.Open "SELECT * FROM users", con, adOpenDynamic, adLockPessimistic
    rs.AddNew
End Sub

Private Sub regbtn_Click()
    rs.Fields("first_name").Value = nametext.Text
    rs.Fields("last_name").Value = lastnametext.Text
    rs.Fields("user_name").Value = usernametext.Text
    rs.Fields("user_password").Value = passwordtext.Text
    rs.Update
    MsgBox "Usuario registrado exitosamente. Por favor inicie sesión", vbInformation
    loginform.Show
    registerform.Hide
End Sub

Private Sub resetbtn_Click()
    nametext.Text = ""
    lastnametext.Text = ""
    usernametext.Text = ""
    passwordtext.Text = ""
End Sub
