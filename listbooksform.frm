VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form listbooksform 
   BackColor       =   &H00400000&
   Caption         =   "Lista de libros"
   ClientHeight    =   7350
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
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
      Height          =   495
      Left            =   11040
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton dislikebtn 
      Caption         =   "No me gustó :("
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10200
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton readbooksbtn 
      Caption         =   "Añadir a leídos"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10200
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton favoritebtn 
      Caption         =   "Agregar a favoritos"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10200
      TabIndex        =   7
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton readbtn 
      Caption         =   "Agregar a leídos"
      Height          =   375
      Left            =   12600
      TabIndex        =   6
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton favoritesbtn 
      Caption         =   "Agregar a favoritos"
      Height          =   375
      Left            =   12600
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton searchbtn 
      Caption         =   "Buscar"
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
      Left            =   3360
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox searchtext 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "Eliminar libro"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin MSComctlLib.ListView bookslist 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7223
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label searchlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar por autor o título de libro"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "listbooksform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public rs As New ADODB.Recordset

Private Sub backbtn_Click()
    userinterface.Show
    Unload Me
End Sub

Private Sub bookslist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    bookslist.Sorted = True
    If bookslist.SortOrder = lvwAscending Then
        bookslist.SortOrder = lvwDescending
    Else
        bookslist.SortOrder = lvwAscending
    End If
End Sub



Private Sub deletebtn_Click()
    If bookslist.SelectedItem Is Nothing Then
        MsgBox "Por favor, selecciona un libro", vbExclamation
        Exit Sub
    End If
    
    Dim confirmation As Integer
    confirmation = MsgBox("¿Estás seguro de eliminar el libro seleccionado?", vbYesNo + vbQuestion, "Confirmar eliminación")
    
    If confirmation = vbYes Then
        If con Is Nothing Then
            Connect
        ElseIf con.State = adStateClosed Then
            Connect
        End If
        'delete selected book
        Dim selectedBook As String
        selectedBook = bookslist.SelectedItem.Text
        
        Dim deletequery As String
        deletequery = "DELETE FROM books WHERE Id = " & selectedBook
        con.Execute deletequery
        bookslist.ListItems.Remove (bookslist.SelectedItem.Index)
        MsgBox "Libro eliminado exitosamente", vbInformation
    End If
End Sub

Public Sub addToFavorites(ByVal userId As Integer, ByVal bookid As Integer)
    Call Connect
    
    Dim sqltable As String
    sqltable = "SELECT * FROM favoritebooks WHERE IdUser = " & userId & " AND IdBook = " & bookid
    rs.Open sqltable, con, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        MsgBox "El libro ya está en tus favoritos", vbExclamation
        rs.Close
        Call Disconnect
        Exit Sub
    End If
    
    rs.Close
    Dim sqlinsert As String
    sqlinsert = "INSERT INTO favoritebooks (IdBook, IdUser, date_favorite) VALUES (" & bookid & ", " & userId & ", GETDATE())"
    rs.Open sqlinsert, con, adOpenStatic, adLockReadOnly
    
    MsgBox "El libro se ha agregado a favoritos", vbInformation
    
    Call Disconnect
    Exit Sub
End Sub

Public Sub addToDislike(ByVal userId As Integer, ByVal bookid As Integer)
    Call Connect
    
    Dim sqltable As String
    sqltable = "SELECT * FROM dislikebooks WHERE IdUser = " & userId & " AND IdBook = " & bookid
    rs.Open sqltable, con, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        MsgBox "El libro ya está agregado", vbExclamation
        rs.Close
        Call Disconnect
        Exit Sub
    End If
    
    rs.Close
    Dim sqlinsert As String
    sqlinsert = "INSERT INTO dislikebooks (IdBook, IdUser, date_dislike) VALUES (" & bookid & ", " & userId & ", GETDATE())"
    rs.Open sqlinsert, con, adOpenStatic, adLockReadOnly
    
    MsgBox "El libro se ha agregado a la lista", vbInformation
    
    Call Disconnect
    Exit Sub
End Sub

Private Sub dislikebtn_Click()
    Dim selectedBook As Integer
    Dim actualuser As Integer
    
    If bookslist.SelectedItem Is Nothing Then
        MsgBox "Por favor, selecciona un libro", vbExclamation
        Exit Sub
    End If
    
    selectedBook = bookslist.SelectedItem.Text
    actualuser = GetActualUserId()
    Call addToDislike(actualuser, selectedBook)
End Sub

Private Sub favoritebtn_Click()
    Dim selectedBook As Integer
    Dim actualuser As Integer
    
    If bookslist.SelectedItem Is Nothing Then
        MsgBox "Por favor, selecciona un libro", vbExclamation
        Exit Sub
    End If
    
    selectedBook = bookslist.SelectedItem.Text
    actualuser = GetActualUserId()
    Call addToFavorites(actualuser, selectedBook)
End Sub
Public Function GetActualUserId() As Integer
    GetActualUserId = actualuser
End Function


    
Private Sub Form_Load()
    With bookslist.ColumnHeaders
        .Add , , "Id", Width / 7, IvwColumnLeft
        .Add , , "Título", Width / 7, IvwColumnCenter
        .Add , , "Autor", Width / 7, IvwColumnCenter
        .Add , , "Género", Width / 7, IvwColumnCenter
        .Add , , "Fecha", Width / 7, IvwColumnCenter
        .Add , , "Sinópsis", Width / 7, IvwColumnCenter
        .Add , , "Imagen", Width / 7, IvwColumnCenter
    End With
loaddata
End Sub

Sub loaddata()
    Dim list As ListItem
    bookslist.ListItems.clear
    
    If con Is Nothing Then
        Connect
    ElseIf con.State = adStateClosed Then
        Connect
    End If
  
    rs.Open "SELECT * FROM books", con, adOpenDynamic, adLockPessimistic
    Do Until rs.EOF
        Set list = bookslist.ListItems.Add(, , rs!Id)
        list.SubItems(1) = rs!Title
        list.SubItems(2) = rs!Author
        list.SubItems(3) = rs!Genre
        list.SubItems(4) = rs!date_book
        list.SubItems(5) = rs!synopsis
        'list.SubItems(6) = rs!Picture
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub readbooksbtn_Click()
    Dim selectedBook As Integer
    Dim actualuser As Integer
    
    If bookslist.SelectedItem Is Nothing Then
        MsgBox "Por favor, selecciona un libro", vbExclamation
        Exit Sub
    End If
    
    selectedBook = bookslist.SelectedItem.Text
    actualuser = GetActualUserId()
    Call addToRead(actualuser, selectedBook)
End Sub

Public Sub addToRead(ByVal userId As Integer, ByVal bookid As Integer)
    Call Connect
    
    Dim sqltable As String
    sqltable = "SELECT * FROM readbooks WHERE IdUser = " & userId & " AND IdBook = " & bookid
    rs.Open sqltable, con, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        MsgBox "EL libro ya está en tus leídos", vbExclamation
        rs.Close
        Call Disconnect
        Exit Sub
    End If
    
    rs.Close
    Dim sqlinsert As String
    sqlinsert = "INSERT INTO readbooks (IdBook, IdUser, date_read) VALUES (" & bookid & ", " & userId & ", GETDATE())"
    rs.Open sqlinsert, con, adOpenStatic, adLockReadOnly
    
    MsgBox "El libro se ha agregado a tus leídos", vbInformation
    
    Call Disconnect
    Exit Sub
End Sub

Private Sub searchbtn_Click()
    Dim itm As ListItem
    Set itm = bookslist.FindItem(searchtext.Text, IvwText, , IvwPartial)
        If itm Is Nothing Then
            MsgBox "Libro no encontrado", vbCritical
        Else
        bookslist.ListItems(itm.Index).Selected = True
        bookslist.SetFocus
        End If
End Sub

Private Sub searchtext_Change()
    bookslist.ListItems.clear
    If con Is Nothing Then
        Connect
    ElseIf con.State = adStateClosed Then
        Connect
    End If
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM books WHERE Author like '" & searchtext.Text & "%' OR Title like '" & searchtext.Text & "%'", con, adOpenDynamic, adLockPessimistic
    Do Until rs.EOF
        Set list = bookslist.ListItems.Add(, , rs!Id)
        list.SubItems(1) = rs!Title
        list.SubItems(2) = rs!Author
        list.SubItems(3) = rs!Genre
        list.SubItems(4) = rs!date_book
        list.SubItems(5) = rs!synopsis
        'list.SubItems(6) = rs!Picture
        rs.MoveNext
    Loop
    Set rs = Nothing
    con.Close: Set con = Nothing
End Sub

