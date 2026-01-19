VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "Calcular"
      Height          =   525
      Left            =   3045
      TabIndex        =   2
      Top             =   3180
      Width           =   2025
   End
   Begin VB.TextBox txtVal2 
      Height          =   285
      Left            =   4530
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1110
      Width           =   1470
   End
   Begin VB.TextBox txtVal1 
      Height          =   330
      Left            =   1020
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1095
      Width           =   1470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalcular_Click()

    Dim intVal1 As Integer
    Dim intVal2 As Integer
    
    intVal1 = CInt(txtVal1.Text)
    intVal2 = CInt(txtVal2.Text)
    
    MsgBox intVal1 + intVal2

End Sub
