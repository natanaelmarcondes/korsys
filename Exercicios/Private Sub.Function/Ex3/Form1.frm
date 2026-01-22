VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMostre 
      Caption         =   "Mostre"
      Height          =   795
      Left            =   2715
      TabIndex        =   1
      Top             =   3540
      Width           =   1425
   End
   Begin VB.TextBox txtEscreva 
      Height          =   405
      Left            =   2670
      TabIndex        =   0
      Top             =   1125
      Width           =   1350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMostre_Click()
    
    MostrarTexto txtEscreva.Text
    
End Sub

Private Sub MostrarTexto(strText As String)
    
    MsgBox strText
    
End Sub
