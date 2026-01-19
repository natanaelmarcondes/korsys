VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBase 
      Caption         =   "Command1"
      Height          =   855
      Left            =   4080
      TabIndex        =   0
      Top             =   3720
      Width           =   2010
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBase_Click()
    
    Dim intNota As Integer
    
    
    intNota = InputBox("Insira sua nota", "Avaliação", "Nota")
    
    
    If intNota >= 7 Then
        MsgBox "Aprovado", vbExclamation + vbOKOnly
    ElseIf intNota >= 5 Then
        MsgBox "Recuperação", vbExclamation + vbOKOnly
    Else
        MsgBox "Reprovado", vbExclamation + vbOKOnly
    End If
    
End Sub
