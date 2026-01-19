VERSION 5.00
Begin VB.Form frmEx3 
   Caption         =   "Ex3"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrim 
      Caption         =   "Command1"
      Height          =   870
      Left            =   2460
      TabIndex        =   1
      Top             =   2670
      Width           =   1785
   End
   Begin VB.TextBox txtTrim 
      Height          =   345
      Left            =   2730
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1560
      Width           =   1125
   End
End
Attribute VB_Name = "frmEx3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTrim_Click()
    
    If Len(Trim(txtTrim.Text)) >= 3 Then
        MsgBox "Registrado", vbInformation
    Else
        MsgBox "Necessario ter mais de três letras"
    End If
    
End Sub
