VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBase 
      Caption         =   "Command1"
      Height          =   795
      Left            =   2340
      TabIndex        =   0
      Top             =   1980
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBase_Click()

    Dim strNome As String

    strNome = InputBox("Digite seu nome")

End Sub
