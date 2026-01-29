VERSION 5.00
Begin VB.Form frmAjuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sôbre o Sistema"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4485
   ScaleMode       =   0  'User
   ScaleWidth      =   8280
   Begin VB.TextBox Text2 
      Height          =   780
      Left            =   1095
      TabIndex        =   2
      Top             =   2520
      Width           =   5910
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Left            =   3225
      TabIndex        =   1
      Top             =   1710
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1125
      TabIndex        =   0
      Top             =   585
      Width           =   5820
   End
End
Attribute VB_Name = "frmAjuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Dim intNome As Integer
    
    intNome = 10
    
    MsgBox TypeName(intNome)
                
    'If Not IsNull(Text1.Text) And Text1.Text <> "" Then
    '    Text2.Text = Text1.Text
    'End If
    'MsgBox IsNumeric(Text1.Text)
    
End Sub

Private Sub Form_Load()
        
    CenterFormInMDI Me, True
    
End Sub

