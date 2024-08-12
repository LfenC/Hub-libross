VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form userReadform 
   BackColor       =   &H00400000&
   Caption         =   "Tus libros léidos"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imglistread 
      Left            =   11040
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
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
      Left            =   10440
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin MSComctlLib.ListView userReadlist 
      Height          =   3975
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
   Begin VB.Label readlabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¡Bienvenido a tus libros leídos!"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   8175
   End
End
Attribute VB_Name = "userReadform"
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
    
    With userReadlist.ColumnHeaders
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
    
    Set rs = GetReadUser(actualuser)
    userReadlist.ListItems.clear

    
    Do Until rs.EOF
        Set list = userReadlist.ListItems.Add(, , rs!Id)
        list.SubItems(1) = rs!Title
        list.SubItems(2) = rs!Author
        list.SubItems(3) = rs!Genre
        list.SubItems(4) = rs!date_book
        list.SubItems(5) = rs!synopsis
        
        rs.MoveNext
    Loop
    
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Disconnect
End Sub

Public Function GetReadUser(userId As Integer) As ADODB.Recordset
    Dim sqlquery As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlquery = "SELECT books.Id, books.Title, books.Author, books.Genre, " & "books.date_book, books.synopsis, books.picture " & "FROM readbooks " & "INNER JOIN books ON readbooks.IdBook = books.Id " & "WHERE readbooks.IdUser = " & userId
    If rs.State = adStateOpen Then
        rs.Close
    End If
    rs.Open sqlquery, con, adOpenStatic, adLockReadOnly
    Set GetReadUser = rs
End Function

