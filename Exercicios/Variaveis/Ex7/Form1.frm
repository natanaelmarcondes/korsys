VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   675
      Left            =   2385
      TabIndex        =   0
      Top             =   2490
      Width           =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim strNumero As String
    Dim intNumero2 As Integer
    
    strNumero = "10"
    intNumero2 = 20
    
    MsgBox CInt(strNumero) + intNumero2

End Sub
