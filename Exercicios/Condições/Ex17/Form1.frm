VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   645
      Left            =   3420
      TabIndex        =   0
      Top             =   4155
      Width           =   1530
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSair_Click()
    
    Dim intSair As Integer
    
    intSair = MsgBox("Tem certeza que deseja sair?", vbYesNo + vbInformation)
    
    If intSair = vbYes Then
        End
    Else
    End If
    
End Sub
