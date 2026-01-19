VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVerificar 
      Caption         =   "Command1"
      Height          =   720
      Left            =   2820
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtIdade 
      Height          =   330
      Left            =   2775
      TabIndex        =   0
      Text            =   "Idade"
      Top             =   1725
      Width           =   1395
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
    
    If intIdade >= 60 Then
        MsgBox "Veio"
    ElseIf intIdade >= 18 Then
        MsgBox "Adulto"
    Else
        MsgBox "Criança"
    End If
    
End Sub
