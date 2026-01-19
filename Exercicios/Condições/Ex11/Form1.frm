VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReg 
      Caption         =   "Registrar"
      Height          =   675
      Left            =   2400
      TabIndex        =   1
      Top             =   3600
      Width           =   1605
   End
   Begin VB.TextBox txtNome 
      Height          =   495
      Left            =   1245
      TabIndex        =   0
      Text            =   "Nome"
      Top             =   1320
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReg_Click()

    If txtNome.Text = "" Then
        MsgBox "Tem que colocar um nome", vbCritical
    Else
        MsgBox "Registrado"
    End If

End Sub
