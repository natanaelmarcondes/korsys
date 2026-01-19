VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   2550
      TabIndex        =   0
      Top             =   3255
      Width           =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim strPreco As String
    
    strPreco = "29,9"
        
    MsgBox CDbl(strPreco) + CDbl("15,9")
 
End Sub
