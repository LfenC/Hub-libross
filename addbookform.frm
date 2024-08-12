VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form addbookform 
   BackColor       =   &H00400000&
   Caption         =   "Acciones con libros"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   12570
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
      Left            =   11040
      TabIndex        =   23
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton lastbtn 
      Caption         =   "Último"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   22
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "Siguiente"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   21
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton previousbtn 
      Caption         =   "Anterior"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   20
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton firstbtn 
      Caption         =   "Primero"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   19
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton findbtn 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker datePicker 
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   3600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      Format          =   139460609
      CurrentDate     =   45512
   End
   Begin VB.CommandButton addbtn 
      Caption         =   "Añadir libro"
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
      Left            =   10200
      TabIndex        =   16
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "Guardar libro"
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
      Left            =   10200
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog PictureDialog 
      Left            =   8040
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton uploadbtn 
      Caption         =   "Subir imagen"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   14
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton clearbtn 
      Caption         =   "Limpiar formulario"
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
      Left            =   10080
      TabIndex        =   13
      Top             =   5520
      Width           =   2175
   End
   Begin VB.TextBox synopsistext 
      Height          =   1695
      Left            =   2400
      TabIndex        =   11
      Top             =   4320
      Width           =   3975
   End
   Begin VB.PictureBox bookpicture 
      AutoRedraw      =   -1  'True
      Height          =   4575
      Left            =   6720
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   10
      Top             =   1440
      Width           =   2970
   End
   Begin VB.TextBox genretext 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox titletext 
      Height          =   405
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox authortext 
      Height          =   405
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
      Width           =   3975
   End
   Begin VB.CommandButton editbtn 
      Caption         =   "Editar libro"
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
      Left            =   10200
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "Eliminar libro"
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
      Left            =   10200
      TabIndex        =   1
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label synopsislabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Sinópsis"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label datelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de publicación"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label genrelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Género"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label titlelabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label autorlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Autor"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label books 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Libros "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "addbookform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs As New ADODB.Recordset

Private Sub addbtn_Click()
    rs.AddNew
    clear
End Sub

Private Sub backbtn_Click()
    userinterface.Show
    Unload Me
End Sub

Private Sub findbtn_Click()
    rs.Close
    rs.Open "SELECT * FROM books WHERE Title='" + titletext.Text + "'", con, adOpenDynamic, adLockPessimistic
    If Not rs.EOF Then
        display
        reload
    Else
    MsgBox "Libro no encontrado", vbInformation
    End If
End Sub

Sub reload()
    rs.Close
    rs.Open "SELECT * FROM books", con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub firstbtn_Click()
    rs.MoveFirst
    display
End Sub

Private Sub Form_Load()
    If rs.State = adStateOpen Then
        rs.Close
    End If
    rs.Open "SELECT * FROM books", con, adOpenDynamic, adLockPessimistic
    display
End Sub
Sub clear()
    titletext.Text = ""
    authortext.Text = ""
    genretext.Text = ""
    datePicker.Value = "16/08/2022"
    synopsistext.Text = ""
    Set bookpicture.Picture = Nothing
End Sub
Private Sub clearbtn_Click()
    titletext.Text = ""
    authortext.Text = ""
    genretext.Text = ""
    datePicker.Value = "16/08/2022"
    synopsistext.Text = ""
    Set bookpicture.Picture = Nothing
End Sub

Private Sub deletebtn_Click()
    confirmation = MsgBox("¿Estás seguro de eliminar este libro?", vbYesNo + vbCritical, "Confirmacion de eliminar el libro")
    If confirmation = vbYes Then
        rs.Delete adAffectCurrent
        MsgBox "El libro ha sido eliminado exitosamente", vbInformation, "Message"
        rs.Update
        refreshdata
    Else
        MsgBox "Eliminación de libro cancelada", vbInformation, "Message"
    End If
End Sub
'refresh data in the recordset, so after deleting the current record, the next record will appear
Sub refreshdata()
    rs.Close
    rs.Open "SELECT * FROM books", con, adOpenStatic, adLockPessimistic
    If Not rs.EOF Then
        rs.MoveNext
        display
    Else
        MsgBox "No records found"
    End If
End Sub
'show the data in the form as form loads
Sub display()
    titletext.Text = rs!Title
    authortext.Text = rs!Author
    genretext.Text = rs!Genre
    datePicker.Value = rs!date_book
    synopsistext.Text = rs!synopsis
    
    'convert to something that the picturebox can show from binary data in the databse
    'verfied if there is a picture (i added some that doesnt have one)
    If Not IsNull(rs!Picture) Then
        Dim tempFilePath As String
        Dim fileNumber As Integer
        Dim toByteData() As Byte
        
        tempFilePath = App.Path & "\temp_image.bmp"
        toByteData = rs!Picture
        fileNumber = FreeFile
        Open tempFilePath For Binary As #fileNumber
        Put #fileNumber, , toByteData
        Close #fileNumber
        bookpicture.Picture = LoadPicture(tempFilePath)
        
        Kill tempFilePath
    Else
        Set bookpicture.Picture = Nothing
    End If
End Sub
Private Sub editbtn_Click()
    rs.Fields("Title").Value = titletext.Text
    rs.Fields("Author").Value = authortext.Text
    rs.Fields("Genre").Value = genretext.Text
    rs.Fields("date_book").Value = datePicker.Value
    rs.Fields("synopsis").Value = synopsistext.Text
    
    'Declare variables to convert image to byte array
    Dim toByteData() As Byte
    Dim fileL As Long
    Dim fileNumber As Integer
    Dim tempFilePath As String
    
    'verified that the user uploade an imagen to change it
    If Not bookpicture.Picture Is Nothing Then
        tempFilePath = App.Path & "\tem_image.bmp"
        SavePicture bookpicture.Picture, tempFilePath
        
        fileNumber = FreeFile
        Open tempFilePath For Binary As #fileNumber
        fileL = LOF(fileNumber)
        
        ReDim toByteData(fileL - 1)
        
        Get #fileNumber, , toByteData
        Close #fileNumber
        rs.Fields("picture").Value = toByteData
        
        Kill tempFilePath
    End If
    rs.Update
    MsgBox "Cambios guardados exitosamente", vbInformation, "Message"
 
End Sub

Private Sub lastbtn_Click()
    rs.MoveLast
    display
End Sub

Private Sub nextbtn_Click()
    rs.MoveNext
    If Not rs.EOF Then
        display
    Else
    rs.MoveFirst
        display
    End If
End Sub

Private Sub previousbtn_Click()
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveLast
        display
    Else
        display
    End If
End Sub

Private Sub savebtn_Click()
'Declare variables to convert image to byte array
    Dim toByteData() As Byte
    Dim fileL As Long
    Dim fileNumber As Integer
    Dim tempFilePath As String
    
    rs.Fields("Title").Value = titletext.Text
    rs.Fields("Author").Value = authortext.Text
    rs.Fields("Genre").Value = genretext.Text
    rs.Fields("date_book").Value = datePicker.Value
    rs.Fields("synopsis").Value = synopsistext.Text
    
    'verified that the user uploade an imagen
    If Not bookpicture.Picture Is Nothing Then
        tempFilePath = App.Path & "\tem_image.bmp"
        SavePicture bookpicture.Picture, tempFilePath
        
        fileNumber = FreeFile
        Open tempFilePath For Binary As #fileNumber
        fileL = LOF(fileNumber)
        
        ReDim toByteData(fileL - 1)
        
        Get #fileNumber, , toByteData
        Close #fileNumber
        rs.Fields("picture").Value = toByteData
        
        Kill tempFilePath
    End If
    rs.Update
    MsgBox "Libro guardado exitosamente", vbInformation, "Message"
End Sub

Private Sub uploadbtn_Click()
    'in case there´s an error
    On Error Resume Next
        PictureDialog.Filter = "Jpeg|*.jpg| Bitmap|*.bmp| png|*.png"
        PictureDialog.ShowOpen
       
        If PictureDialog.FileName <> "" Then
            bookpicture.Picture = LoadPicture(PictureDialog.FileName)
        End If
End Sub
