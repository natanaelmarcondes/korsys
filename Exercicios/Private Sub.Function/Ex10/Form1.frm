VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVazio 
      Height          =   495
      Left            =   3345
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1035
      Width           =   1110
   End
   Begin VB.CommandButton cmdChecar 
      Caption         =   "Checar"
      Height          =   780
      Left            =   3285
      TabIndex        =   0
      Top             =   3525
      Width           =   1545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChecar_Click()
    
    ContinuarSub
    
End Sub

Private Sub ContinuarSub()
    
    If Validando() = False Then
        Exit Sub
    End If
    
    MsgBox "continuando o sub"

End Sub


Private Function Validando() As Boolean


    If Trim(txtVazio.Text) = "" Then
        Validando = False
        Exit Function
    End If
    
    Validando = True

End Function
