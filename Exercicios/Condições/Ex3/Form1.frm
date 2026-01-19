VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBase 
      Caption         =   "Command1"
      Height          =   840
      Left            =   2970
      TabIndex        =   0
      Top             =   3240
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBase_Click()
    
    Dim intNumero As Integer
    
    intNumero = InputBox("Insira um numero a ser identificado", "Identificação", 0)
    
    If intNumero > 0 Then
        MsgBox "Seu numero é positivo"
    ElseIf intNumero < 0 Then
        MsgBox "Seu numero é negativo"
    Else
        MsgBox "Seu numero é Zero"
    End If
    
    
End Sub
