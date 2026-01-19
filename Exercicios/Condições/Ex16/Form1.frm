VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClick 
      Caption         =   "Bonito"
      Height          =   1035
      Left            =   2670
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3255
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClick_Click()
    
    MsgBox "Botão bonito"
    
End Sub
