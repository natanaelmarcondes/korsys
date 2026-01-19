VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Command1"
      Height          =   870
      Left            =   2550
      TabIndex        =   0
      Top             =   2415
      Width           =   1590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExibir_Click()

    Dim intIdade As Integer
    Dim strIdade2 As String
    
    strIdade2 = "25"
    
    intIdade = CInt(strIdade2)
    
    
    
    MsgBox intIdade

End Sub
