VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form userFavform 
   BackColor       =   &H00400000&
   Caption         =   "Tus favoritos"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton backbtn 
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin MSComctlLib.ListView userfavlist 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   4194304
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label favlabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¡Bienvenido a tus libros favoritos!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   10335
   End
End
Attribute VB_Name = "userFavform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub backbtn_Click()
    userinterface.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim list As ListItem
    
    With userfavlist.ColumnHeaders
        .Add , , "Id libro", Width / 7, IvwColumnLeft
        .Add , , "Título", Width / 7, IvwColumnCenter
        .Add , , "Autor", Width / 7, IvwColumnCenter
        .Add , , "Género", Width / 7, IvwColumnCenter
        .Add , , "Fecha de publicación", Width / 7, IvwColumnCenter
        .Add , , "Sinópsis", Width / 7, IvwColumnCenter
        .Add , , "Imagen", Width / 7, IvwColumnCenter
    End With
    
    actualuser = GetActualUserId()
    
    Connect
    
    Set rs = GetFavoritesUser(actualuser)
    userfavlist.ListItems.clear
    
    Do Until rs.EOF
        Set list = userfavlist.ListItems.Add(, , rs!Id)
        list.SubItems(1) = rs!Title
        list.SubItems(2) = rs!Author
        list.SubItems(3) = rs!Genre
        list.SubItems(4) = rs!date_book
        list.SubItems(5) = rs!synopsis
        'list.SubItems(6) = rs!Picture
        rs.MoveNext
    Loop
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Disconnect
End Sub

Public Function GetFavoritesUser(userId As Integer) As ADODB.Recordset
    Dim sqlquery As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlquery = "SELECT books.Id, books.Title, books.Author, books.Genre, " & "books.date_book, books.synopsis " & "FROM favoritebooks " & "INNER JOIN books ON favoritebooks.IdBook = books.Id " & "WHERE favoritebooks.IdUser = " & userId
    If rs.State = adStateOpen Then
        rs.Close
    End If
    rs.Open sqlquery, con, adOpenStatic, adLockReadOnly
    Set GetFavoritesUser = rs
End Function

