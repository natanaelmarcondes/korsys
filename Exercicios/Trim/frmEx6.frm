VERSION 5.00
Begin VB.Form frmEx6 
   Caption         =   "Ex6"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdJuntar 
      Caption         =   "Juntar três letras iniciais e finais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   2565
      TabIndex        =   2
      Top             =   3180
      Width           =   2760
   End
   Begin VB.TextBox txtLeftRight 
      Height          =   510
      Left            =   2955
      TabIndex        =   0
      Top             =   825
      Width           =   1395
   End
   Begin VB.Label lblLeftRight 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2445
      TabIndex        =   1
      Top             =   1890
      Width           =   2475
   End
End
Attribute VB_Name = "frmEx6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdJuntar_Click()
    
    strLeftRight = txtLeftRight.Text
    
    
    lblLeftRight.Caption = Left(strLeftRight, 3) & "< >" & Right(strLeftRight, 2)
    
    
End Sub
