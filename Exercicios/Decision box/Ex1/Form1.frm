VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBase 
      Caption         =   "Command1"
      Height          =   660
      Left            =   3750
      TabIndex        =   0
      Top             =   3030
      Width           =   1740
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBase_Click()

    MsgBox "Bem vindo- ao curso de programação VB6", vbOKOnly
    
End Sub
