VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   2640
      Width           =   1980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim intTotal As Integer
    Dim strMensagem As String
    
    intTotal = 25
    strMensagem = "Robson"
    
    MsgBox CStr(intTotal) & " " & strMensagem

End Sub
