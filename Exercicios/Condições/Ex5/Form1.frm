VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBase 
      Caption         =   "Command1"
      Height          =   1080
      Left            =   3570
      TabIndex        =   0
      Top             =   3660
      Width           =   2640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBase_Click()
    
    Dim intNumero1 As Integer
    Dim intNumero2 As Integer
    
    intNumero1 = Val(InputBox("Numero 1"))
    intNumero2 = Val(InputBox("Numero 2"))
    
    If intNumero1 = intNumero2 Then
        MsgBox "Os numeros são iguais"
    ElseIf intNumero1 < intNumero2 Then
        MsgBox "O numero " & intNumero2 & " é maior"
    Else
        MsgBox "O numero " & intNumero1 & " é maior"
    End If
    
    
    
    
End Sub
