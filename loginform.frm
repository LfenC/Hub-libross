VERSION 5.00
Begin VB.Form loginform 
   BackColor       =   &H00400000&
   Caption         =   "Login form"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cancelbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   6
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton loginbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ingresar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox passwordtext 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox usernametext 
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label welcomelabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¡Bienvenido a tu hub de libros!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label passwordlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label usernamelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de usuario"
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
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label nameloginlabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciar sesión"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "loginform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs As New ADODB.Recordset
Public idUserA As Integer


Private Sub cancelbtn_Click()
    welcomeform.Show
    Unload Me
End Sub

Private Sub loginbtn_Click()
    If con Is Nothing Then
        Connect
    ElseIf con.State = adStateClosed Then
        Connect
    End If
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    rs.Open "SELECT * FROM users WHERE user_name='" & usernametext.Text & "' AND user_password ='" & passwordtext.Text & "'", con, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then
        MsgBox "Error al iniciar sesión. Por favor inicia sesión con datos correctos", vbCritical
        loginform.Show
        usernametext.Text = ""
        passwordtext.Text = ""
    Else
        actualuser = rs!Id
        MsgBox "Inicio de sesión exitoso", vbInformation
        userinterface.Show
        loginform.Hide
    End If
    rs.Close
End Sub


