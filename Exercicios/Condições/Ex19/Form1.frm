VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChecar 
      Caption         =   "Command1"
      Height          =   900
      Left            =   2145
      TabIndex        =   2
      Top             =   3885
      Width           =   2325
   End
   Begin VB.TextBox txtIdade 
      Height          =   405
      Left            =   3630
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1365
      Width           =   1200
   End
   Begin VB.TextBox txtNome 
      Height          =   510
      Left            =   1215
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1305
      Width           =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChecar_Click()
    
    
    If (txtNome.Text = "") Or (txtIdade.Text = "") Then
        MsgBox "é necessario preencher algum dos dois ai em cima", vbCritical
    Else
        MsgBox "Olá " & txtNome.Text & " de " & txtIdade.Text & " Anos"
    End If

    
End Sub
