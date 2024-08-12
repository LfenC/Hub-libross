VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form splash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Old Antic Outline"
      Size            =   18
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timersplash 
      Interval        =   250
      Left            =   10200
      Top             =   360
   End
   Begin MSComctlLib.ProgressBar progressBar 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   4920
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
      Max             =   105
      Scrolling       =   1
   End
   Begin VB.Label namelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Hub de libros"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   600
      TabIndex        =   3
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Label labelstat 
      BackStyle       =   0  'Transparent
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
      Left            =   4080
      TabIndex        =   2
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label labelstatus 
      BackStyle       =   0  'Transparent
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
      Left            =   480
      TabIndex        =   1
      Top             =   4440
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF80FF&
      Height          =   5895
      Left            =   0
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    timersplash.Enabled = True
    'adjust the screen to the middle
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub


Private Sub timersplash_Timer()
    progressBar.Value = progressBar.Value + 5
    labelstatus.Caption = "Cargando,  por  favor  espere..."
    labelstat.Caption = progressBar.Value & "%"
    If progressBar.Value = progressBar.Max Then
        timersplash.Enabled = False
        Unload Me
        welcomeform.Show
    End If
End Sub
