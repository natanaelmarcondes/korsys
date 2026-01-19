VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4845
      Top             =   3510
   End
   Begin VB.Label lblHora 
      Caption         =   "Label1"
      Height          =   570
      Left            =   2310
      TabIndex        =   0
      Top             =   2790
      Width           =   1410
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()

    lblHora.Caption = Time
     
End Sub
