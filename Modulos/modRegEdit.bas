Attribute VB_Name = "modRegEdit"
Option Explicit
' ============================
' Regedit simples (VB6)
' Usa HKCU\Software\VB and VBA Program Settings
' ============================

Public Sub GReg_Gravar(ByVal PStr_App As String, _
                       ByVal PStr_Secao As String, _
                       ByVal PStr_Chave As String, _
                       ByVal PStr_Valor As String)
    On Error GoTo Erro
    SaveSetting PStr_App, PStr_Secao, PStr_Chave, PStr_Valor
    Exit Sub
Erro:
    MsgBox "Erro ao gravar no Registro: " & Err.Description, vbCritical
End Sub

Public Function GReg_Ler(ByVal PStr_App As String, _
                         ByVal PStr_Secao As String, _
                         ByVal PStr_Chave As String, _
                         Optional ByVal PStr_Padrao As String = "") As String
    On Error GoTo Erro
    GReg_Ler = GetSetting(PStr_App, PStr_Secao, PStr_Chave, PStr_Padrao)
    Exit Function
Erro:
    MsgBox "Erro ao ler do Registro: " & Err.Description, vbCritical
    GReg_Ler = PStr_Padrao
End Function

