VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExecutar 
      Caption         =   "Command1"
      Height          =   690
      Left            =   3645
      TabIndex        =   0
      Top             =   3540
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExecutar_Click()
    
    Dim strNome(0 To 4) As String
    Dim i As Integer
    
    strNome(0) = "Teste1"
    strNome(1) = ""
    strNome(2) = "Teste3"
    strNome(3) = ""
    strNome(4) = "Teste5"

    For i = 0 To UBound(strNome) 'Step 2
        MsgBox strNome(i)
    Next


End Sub
