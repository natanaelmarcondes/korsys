VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Menu"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEx10 
      Caption         =   "Ex10"
      Height          =   675
      Left            =   3330
      TabIndex        =   9
      Top             =   4890
      Width           =   1695
   End
   Begin VB.CommandButton cmdEx9 
      Caption         =   "Ex9"
      Height          =   795
      Left            =   630
      TabIndex        =   8
      Top             =   4905
      Width           =   1620
   End
   Begin VB.CommandButton cmdEx8 
      Caption         =   "Ex8"
      Height          =   765
      Left            =   9000
      TabIndex        =   7
      Top             =   3045
      Width           =   1770
   End
   Begin VB.CommandButton cmdEx7 
      Caption         =   "Ex7"
      Height          =   765
      Left            =   6150
      TabIndex        =   6
      Top             =   2925
      Width           =   1770
   End
   Begin VB.CommandButton cmdEx6 
      Caption         =   "Ex6"
      Height          =   780
      Left            =   3315
      TabIndex        =   5
      Top             =   2940
      Width           =   1770
   End
   Begin VB.CommandButton cmdEx5 
      Caption         =   "Ex5"
      Height          =   930
      Left            =   585
      TabIndex        =   4
      Top             =   2925
      Width           =   1725
   End
   Begin VB.CommandButton cmdEx4 
      Caption         =   "Ex4"
      Height          =   750
      Left            =   8880
      TabIndex        =   3
      Top             =   795
      Width           =   1680
   End
   Begin VB.CommandButton cmdEx3 
      Caption         =   "Ex3"
      Height          =   750
      Left            =   6015
      TabIndex        =   2
      Top             =   750
      Width           =   1860
   End
   Begin VB.CommandButton cmdEx2 
      Caption         =   "Ex2"
      Height          =   810
      Left            =   3300
      TabIndex        =   1
      Top             =   735
      Width           =   1755
   End
   Begin VB.CommandButton cmdEx1 
      Caption         =   "Ex1"
      Height          =   795
      Left            =   540
      TabIndex        =   0
      Top             =   765
      Width           =   1710
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEx1_Click()

    frmEx1.Show

End Sub

Private Sub cmdEx10_Click()
    
    frmEx10.Show
    
End Sub

Private Sub cmdEx2_Click()

    frmEx2.Show

End Sub

Private Sub cmdEx3_Click()
    
    frmEx3.Show
    
End Sub

Private Sub cmdEx4_Click()
    
    frmEx4.Show
    
End Sub

Private Sub cmdEx5_Click()
    
    frmEx5.Show
    
End Sub

Private Sub cmdEx6_Click()
    
    frmEx6.Show
    
End Sub

Private Sub cmdEx7_Click()
    
    frmEx7.Show
    
End Sub

Private Sub cmdEx8_Click()

    frmEx8.Show

End Sub

Private Sub cmdEx9_Click()

    frmEx9.Show

End Sub
