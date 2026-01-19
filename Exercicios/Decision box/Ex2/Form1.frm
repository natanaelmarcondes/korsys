VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBase 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2955
      TabIndex        =   0
      Top             =   3420
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBase_Click()

    MsgBox "Cadastro concluído com sucesso", vbInformation, "Sistema"

End Sub
