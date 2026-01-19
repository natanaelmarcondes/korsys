VERSION 5.00
Begin VB.Form frmEx2 
   Caption         =   "Ex2"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   465
      Left            =   5565
      TabIndex        =   3
      Top             =   2550
      Width           =   945
   End
   Begin VB.CommandButton cmdLoggedUser 
      Caption         =   "Checar Conta"
      Height          =   660
      Left            =   2835
      TabIndex        =   2
      Top             =   3510
      Width           =   1875
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registro"
      Height          =   735
      Left            =   2835
      TabIndex        =   0
      Top             =   945
      Width           =   1920
   End
   Begin VB.Label lblUsuario 
      Height          =   735
      Left            =   2535
      TabIndex        =   1
      Top             =   2100
      Width           =   2505
   End
End
Attribute VB_Name = "frmEx2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdLimpar_Click()

    lblUsuario.Caption = Trim("")

End Sub

Private Sub cmdLoggedUser_Click()

    lblUsuario.Caption = strUsuario

End Sub

Private Sub cmdRegistrar_Click()

    strUsuario = InputBox("Escreva seu nome")
    
    lblUsuario.Caption = strUsuario
    
End Sub

