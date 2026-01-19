VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBase 
      Caption         =   "Command1"
      Height          =   690
      Left            =   2865
      TabIndex        =   0
      Top             =   3000
      Width           =   1905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBase_Click()

    Dim intConfirmacao As Integer

    intConfirmacao = MsgBox("Confirma a operação?", vbExclamation + vbOKCancel, "Confirmação")

End Sub
