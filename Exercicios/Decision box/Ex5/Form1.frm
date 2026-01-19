VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBase 
      Caption         =   "Command1"
      Height          =   780
      Left            =   2385
      TabIndex        =   0
      Top             =   2370
      Width           =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBase_Click()

    Dim intEscolha As Integer
    
    intEscolha = MsgBox("Deseja continuar?", vbQuestion + vbYesNo)
    


End Sub
