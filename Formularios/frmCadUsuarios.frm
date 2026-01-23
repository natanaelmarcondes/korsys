VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCadUsuarios 
   Caption         =   "frmCadUsuarios"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   7905
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1140
      Top             =   3015
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Height          =   450
      Left            =   6240
      TabIndex        =   0
      Top             =   4230
      Width           =   1515
   End
End
Attribute VB_Name = "frmCadUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

