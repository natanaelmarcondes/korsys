VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   720
      Left            =   2625
      TabIndex        =   0
      Top             =   3345
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim strProduto As String
    Dim intQtd As Integer
    Dim curPreco As Currency
    
    strProduto = "Teclado"
    intQtd = 4
    curPreco = "159,99"
    
    MsgBox "Produto: " & strProduto & " Quantidade: " & intQtd & " Preço: " & curPreco
    

End Sub
