VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   960
      Left            =   2790
      TabIndex        =   1
      Top             =   3420
      Width           =   2025
   End
   Begin VB.TextBox txtSaldo 
      Height          =   450
      Left            =   1905
      TabIndex        =   0
      Text            =   "Saldo"
      Top             =   1485
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Dim intSaldo As Integer
    
    intSaldo = Val(txtSaldo.Text)
    
    If intSaldo = 0 Then
        MsgBox "Saldo zerado"
    ElseIf intSaldo > 0 Then
        MsgBox "Saldo Positivo"
    Else
        MsgBox "Saldo Negativo"
    End If
    
End Sub
