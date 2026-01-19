VERSION 5.00
Begin VB.Form frmEx4 
   Caption         =   "Ex4"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMinusc 
      Caption         =   "Minusculo"
      Height          =   720
      Left            =   4890
      TabIndex        =   2
      Top             =   3525
      Width           =   1440
   End
   Begin VB.CommandButton cmdMaiusc 
      Caption         =   "Maiusculo"
      Height          =   765
      Left            =   1770
      TabIndex        =   1
      Top             =   3510
      Width           =   1560
   End
   Begin VB.TextBox txtMaiMin 
      Height          =   465
      Left            =   3405
      TabIndex        =   0
      Text            =   "Texto de Teste"
      Top             =   960
      Width           =   1800
   End
End
Attribute VB_Name = "frmEx4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMaiusc_Click()
    
    txtMaiMin.Text = UCase(txtMaiMin.Text)
    
End Sub

Private Sub cmdMinusc_Click()
    
    txtMaiMin.Text = LCase(txtMaiMin.Text)
    
End Sub
