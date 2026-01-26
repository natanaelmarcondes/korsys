VERSION 5.00
Begin VB.Form frmConfiguracoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurações do Sistema"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7005
   Icon            =   "frmConfiguracoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7005
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      Height          =   615
      Left            =   -45
      ScaleHeight     =   555
      ScaleWidth      =   7080
      TabIndex        =   8
      Top             =   -15
      Width           =   7140
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Configurações"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   180
         TabIndex        =   9
         Top             =   75
         Width           =   4425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configurações"
      Height          =   2475
      Left            =   150
      TabIndex        =   5
      Top             =   780
      Width           =   6645
      Begin VB.TextBox txtValor 
         Height          =   330
         Index           =   1
         Left            =   165
         TabIndex        =   1
         ToolTipText     =   "Digite o Endereço da API dos correios"
         Top             =   1140
         Width           =   6255
      End
      Begin VB.TextBox txtValor 
         Height          =   330
         Index           =   0
         Left            =   150
         TabIndex        =   0
         ToolTipText     =   "Digite o endereço da API do Discord"
         Top             =   555
         Width           =   6255
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         Height          =   630
         Left            =   5445
         TabIndex        =   2
         Top             =   1740
         Width           =   1020
      End
      Begin VB.Label Label2 
         Caption         =   "Endereço Cep"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   7
         Top             =   915
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "Endereço Discord"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   180
         TabIndex        =   6
         Top             =   285
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   825
      Left            =   5085
      Picture         =   "frmConfiguracoes.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3345
      Width           =   1710
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3285
      Picture         =   "frmConfiguracoes.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3345
      Width           =   1815
   End
End
Attribute VB_Name = "frmConfiguracoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdSalvar_Click()
    
    GravarINI "WEBHOOK", "DISCORD", txtValor(0).Text
    GravarINI "WEBHOOK", "CORREIOS", txtValor(1).Text
        
End Sub

'https://discord.com/api/webhooks/1464374438358155548/AodvRxiM6Jcp50lDRVHKEcKfWZRgoQhbgLk2CvCHY7IrKsChTkPB0wjsF4-7WkYZRzRL
Private Sub Form_Load()
    
    Me.Height = 4725
    Me.Width = 7170
    
    txtValor(0).Text = LerINI("WEBHOOK", "DISCORD")
    txtValor(1).Text = LerINI("WEBHOOK", "CORREIOS")
    
    
    
End Sub
