VERSION 5.00
Begin VB.Form frmInicio 
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFoco 
      Caption         =   "Foco"
      Height          =   495
      Left            =   3735
      TabIndex        =   1
      Top             =   2385
      Width           =   1605
   End
   Begin VB.TextBox txtNome 
      Height          =   555
      Left            =   3555
      TabIndex        =   0
      Text            =   "Teste de foco"
      Top             =   3675
      Width           =   1920
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFoco_Click()

    MsgBox "esse botão é só pra ter algum outro lugar para tirar o foco"

End Sub


Private Sub txtNome_GotFocus()
    
    txtNome.BackColor = vbBlue
    frmInicio.BackColor = vbBlue
    
End Sub

Private Sub txtNome_LostFocus()
    
    txtNome.BackColor = vbWhite
    frmInicio.BackColor = vbWhite
    
End Sub
