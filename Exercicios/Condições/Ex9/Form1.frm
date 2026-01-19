VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Command1"
      Height          =   795
      Left            =   2895
      TabIndex        =   1
      Top             =   3270
      Width           =   1560
   End
   Begin VB.CheckBox chkBox 
      Caption         =   "Ativo/Inativo"
      Height          =   555
      Left            =   1770
      TabIndex        =   0
      Top             =   1725
      Width           =   2070
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
    
    If chkBox.Value = 1 Then
        MsgBox "Ativo"
    Else
        MsgBox "Inativo"
    End If
    
End Sub
