VERSION 5.00
Begin VB.Form frmEx5 
   Caption         =   "Ex5"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdValor 
      Caption         =   "É numerico?"
      Height          =   870
      Left            =   2940
      TabIndex        =   1
      Top             =   3735
      Width           =   1905
   End
   Begin VB.TextBox txtVal 
      Height          =   435
      Left            =   2970
      TabIndex        =   0
      Top             =   1905
      Width           =   1530
   End
End
Attribute VB_Name = "frmEx5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdValor_Click()
    
    strIsNumeric = txtVal.Text
    
    If IsNumeric(strIsNumeric) Then
        MsgBox "é um numero"
    Else
        MsgBox "não é um numero", vbCritical
        txtVal.SetFocus
        txtVal.Text = ""
    End If
    
End Sub
