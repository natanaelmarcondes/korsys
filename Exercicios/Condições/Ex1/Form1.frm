VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVerificar 
      Caption         =   "Verificar"
      Height          =   510
      Left            =   3180
      TabIndex        =   1
      Top             =   3795
      Width           =   1950
   End
   Begin VB.TextBox txtIdade 
      Height          =   375
      Left            =   3375
      TabIndex        =   0
      Top             =   1125
      Width           =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVerificar_Click()

    Dim intIdade As Integer
    
    intIdade = Val(txtIdade.Text)
    
    If intIdade >= 18 Then
        MsgBox "Você é maior de idade"
    Else
        MsgBox "Você é menor de idade"
    End If

End Sub
