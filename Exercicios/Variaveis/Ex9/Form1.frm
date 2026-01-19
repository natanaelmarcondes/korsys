VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   690
      Left            =   2910
      TabIndex        =   0
      Top             =   2865
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim strQtd As String
    Dim strPreco As String
    
    strQtd = "154,12"
    strPreco = "263,54"
    
    MsgBox CCur(strQtd) + CCur(strPreco)

End Sub
