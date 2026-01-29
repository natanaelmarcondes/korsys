VERSION 5.00
Begin VB.Form frmAjuda 
   Caption         =   "Sôbre o Sistema"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleMode       =   0  'User
   ScaleWidth      =   8385
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   285
      Index           =   2
      Left            =   150
      TabIndex        =   7
      Top             =   1065
      Width           =   270
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   285
      Index           =   1
      Left            =   135
      TabIndex        =   6
      Top             =   615
      Width           =   270
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   285
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   195
      Value           =   -1  'True
      Width           =   270
   End
   Begin VB.TextBox Text2 
      Height          =   780
      Left            =   1890
      TabIndex        =   2
      Top             =   4680
      Width           =   5910
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Left            =   4200
      TabIndex        =   1
      Top             =   4020
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1905
      TabIndex        =   0
      Top             =   3255
      Width           =   5820
   End
   Begin VB.Label Label2 
      Caption         =   "2- left e right"
      Height          =   330
      Left            =   510
      TabIndex        =   4
      Top             =   630
      Width           =   3225
   End
   Begin VB.Label Label1 
      Caption         =   "1 - mid"
      Height          =   330
      Left            =   555
      TabIndex        =   3
      Top             =   180
      Width           =   3225
   End
End
Attribute VB_Name = "frmAjuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Select Case True
        
        Case Option1(0).Value
            
            
            'Text2.Text = Len(Text1.Text)
            
            Text2.Text = Replace(Text1.Text, ",", ".")
            
        Case Option1(1).Value
                    
            

    End Select
    
End Sub

Private Sub Form_Load()
    CenterFormInMDI Me
End Sub

