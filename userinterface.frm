VERSION 5.00
Begin VB.Form userinterface 
   BackColor       =   &H00400000&
   Caption         =   "Mis libros"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8220
   FillColor       =   &H00C000C0&
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exitbtn 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   8
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton booksbtn 
      Caption         =   "Añadir libro"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   6
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton searchbtn 
      BackColor       =   &H8000000D&
      Caption         =   "Ver libros "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      MaskColor       =   &H0000FF00&
      TabIndex        =   5
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton dislikebtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Libros que no me gustaron"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5040
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton readbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ver mis libros leídos"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   3
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton favgenresbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ver mis géneros favoritos"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton favoritesbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ver mis favoritos"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label infolabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona una opción"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label yourbooks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tus libros "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "userinterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub booksbtn_Click()
    Call Connect
    Dim formBooks As New addbookform
    formBooks.Show
    Me.Hide
End Sub

Private Sub dislikebtn_Click()
    Call Connect
    Dim mydislikes As New dislikeform
    mydislikes.Show
    Me.Hide
End Sub

Private Sub exitbtn_Click()
    End
End Sub

Private Sub favoritesbtn_Click()
    Call Connect
    Dim myfavorites As New userFavform
    myfavorites.Show
    Me.Hide
End Sub

Private Sub readbtn_Click()
    Call Connect
    Dim myreadbooks As New userReadform
    myreadbooks.Show
    Me.Hide
End Sub

Private Sub searchbtn_Click()
    Call Connect
    Dim listbook As New listbooksform
    listbook.Show
    Me.Hide
End Sub
