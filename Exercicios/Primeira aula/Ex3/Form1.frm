VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEscreva 
      Height          =   1065
      Left            =   330
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2655
      Width           =   3510
   End
   Begin VB.Label lblTextoEscreva 
      Caption         =   "Label1"
      Height          =   360
      Left            =   4125
      TabIndex        =   1
      Top             =   1335
      Width           =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtEscreva_Change()

    lblTextoEscreva.Caption = txtEscreva.Text

End Sub
