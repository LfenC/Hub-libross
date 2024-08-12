Attribute VB_Name = "db_module"
Public actualuser As Integer
Option Explicit
Public con As ADODB.Connection
Public rs As New ADODB.Recordset

Public Sub Connect()
    Set con = New ADODB.Connection
    con.ConnectionString = "Provider= SQLNCLI11; Data Source=LAPTOP-OIMS98C3; Initial Catalog=booksdb; User ID=lizethcbr; Password=Lcfibonacci0501;"
    con.Open
End Sub

Public Sub Disconnect()
    If Not con Is Nothing Then
        con.Close
        Set con = Nothing
    End If
End Sub

Public Function GetActualUserId() As Integer
    GetActualUserId = actualuser
End Function

