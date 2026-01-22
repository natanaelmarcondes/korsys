VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdValidar 
      Caption         =   "Validar"
      Height          =   990
      Left            =   2820
      TabIndex        =   1
      Top             =   3840
      Width           =   1995
   End
   Begin VB.TextBox txtIdade 
      Height          =   405
      Left            =   3585
      TabIndex        =   0
      Top             =   1695
      Width           =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdValidar_Click()

    Dim iIdade As Integer
    iIdade = Val(txtIdade.Text)

    If ValidarIdade(iIdade) Then
        MsgBox "true"
    Else
        MsgBox "false!"
    End If

End Sub


Private Function ValidarIdade(intIdade As Integer) As Boolean
    
    ValidarIdade = (intIdade >= 3 And intIdade <= 18)
    
End Function
