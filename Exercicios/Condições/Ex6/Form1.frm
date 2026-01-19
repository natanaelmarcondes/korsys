VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2595
      TabIndex        =   1
      Top             =   3420
      Width           =   1800
   End
   Begin VB.TextBox txtValor 
      Height          =   465
      Left            =   2310
      TabIndex        =   0
      Text            =   "Valor"
      Top             =   1290
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Dim intValor As Integer
    
    intValor = Val(txtValor.Text)
    
    If intValor >= 500 Then
        MsgBox "Desconto especial"
    Else
        MsgBox "sem desconto pra compra barata"
    End If

End Sub
