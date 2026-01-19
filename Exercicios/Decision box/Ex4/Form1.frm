VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBase 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2235
      TabIndex        =   0
      Top             =   1845
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBase_Click()

    Dim strCidade As String
    
    strCidade = InputBox("Digite sua cidade", "Cadastro", "São Paulo")
    

End Sub
