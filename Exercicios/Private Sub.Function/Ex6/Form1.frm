VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRetorno 
      Caption         =   "Retornar"
      Height          =   870
      Left            =   3255
      TabIndex        =   0
      Top             =   3570
      Width           =   1605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRetorno_Click()
    
    MsgBox AnoAtual
    
End Sub
    
Private Function AnoAtual() As Integer
    
    AnoAtual = Year(Date)
    
End Function
