VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextBoxSimples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================
' Variáveis internas
'==============================
Private m_Text As String
Private m_MaxLength As Long
Private m_Enabled As Boolean

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal Value As String)
    If m_Enabled = False Then Exit Property
    
    If m_MaxLength > 0 Then
        m_Text = Left$(Value, m_MaxLength)
    Else
        m_Text = Value
    End If
End Property
Public Property Get MaxLength() As Long
    MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(ByVal Value As Long)
    If Value < 0 Then Value = 0
    m_MaxLength = Value
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    m_Enabled = Value
End Property

Public Sub Clear()
    m_Text = ""
End Sub

Private Sub Class_Initialize()
    m_Text = ""
    m_MaxLength = 0
    m_Enabled = True
End Sub
