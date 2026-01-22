VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSomar 
      Caption         =   "Somar"
      Height          =   720
      Left            =   2550
      TabIndex        =   2
      Top             =   4065
      Width           =   2220
   End
   Begin VB.TextBox txtNum2 
      Height          =   285
      Left            =   3990
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1845
      Width           =   1095
   End
   Begin VB.TextBox txtNum1 
      Height          =   345
      Left            =   2205
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1830
      Width           =   1185
   End
   Begin VB.Label lblResultado 
      Caption         =   "Label1"
      Height          =   495
      Left            =   6405
      TabIndex        =   3
      Top             =   2805
      Width           =   1350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function SomarNumeros(dNum1 As Double, dNum2 As Double) As Double
    
    SomarNumeros = dNum1 + dNum2
    
End Function

Private Sub cmdSomar_Click()
    
    
    lblResultado.Caption = SomarNumeros(CDbl(txtNum1), CDbl(txtNum2))
    
    
End Sub
