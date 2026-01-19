VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Command1"
      Height          =   1005
      Left            =   1935
      TabIndex        =   2
      Top             =   3120
      Width           =   2010
   End
   Begin VB.TextBox txtSujo 
      Height          =   465
      Left            =   1140
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1170
      Width           =   1065
   End
   Begin VB.Label lblSujo 
      Caption         =   "Label1"
      Height          =   765
      Left            =   3465
      TabIndex        =   1
      Top             =   1350
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLimpar_Click()
    
    txtSujo.Text = ""
    lblSujo.Caption = ""
    
End Sub
