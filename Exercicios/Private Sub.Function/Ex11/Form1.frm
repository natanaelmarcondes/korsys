VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTeste 
      Caption         =   "Teste"
      Height          =   900
      Left            =   3300
      TabIndex        =   1
      Top             =   3675
      Width           =   2040
   End
   Begin VB.TextBox txtTeste 
      Height          =   660
      Left            =   3480
      TabIndex        =   0
      Top             =   1440
      Width           =   1545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTeste_Click()
    Dim strTeste As String
    
    strTeste = txtTeste.Text
    
    If TesteFunct(strTeste) Then
        TesteResultado CDbl(strTeste) + NumeroBase()
    End If
    
End Sub

Private Sub TesteResultado(dblNum As Double)
    
    MsgBox "Numero: " & dblNum
    
End Sub


Private Function NumeroBase() As Double

    NumeroBase = 44
    
End Function


Private Function TesteFunct(strTeste As String) As Boolean
    
    TesteFunct = IsNumeric(strTeste)
    
End Function
