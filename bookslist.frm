VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form bookslist 
   Caption         =   "Lista de libros"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6165
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "bookslist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
