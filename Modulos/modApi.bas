Attribute VB_Name = "modApi"
Option Explicit

'Beep
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

'Sleep / Pausa
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'========================================================
' INI via Windows API
'========================================================
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
Public Function LerINI(ByVal PPStr_Secao As String, _
                       ByVal PPStr_Chave As String, _
                       Optional ByVal PPStr_Padrao As String = "", _
                       Optional ByVal PPStr_Arquivo As String = "") As String
    Dim WLStr_Buffer As String
    Dim WLLng_Tam As Long
    Dim WLStr_IniPath As String
    
    WLStr_IniPath = IIf(Trim$(PPStr_Arquivo) <> "", PPStr_Arquivo, App.Path & "\KORSYS.ini")
    
    WLStr_Buffer = String$(2048, vbNullChar)
    WLLng_Tam = GetPrivateProfileString(PPStr_Secao, PPStr_Chave, PPStr_Padrao, WLStr_Buffer, Len(WLStr_Buffer), WLStr_IniPath)
    
    LerINI = Left$(WLStr_Buffer, WLLng_Tam)
End Function

Public Function GravarINI(ByVal PPStr_Secao As String, _
                          ByVal PPStr_Chave As String, _
                          ByVal PPStr_Valor As String, _
                          Optional ByVal PPStr_Arquivo As String = "") As Boolean
    Dim WLStr_IniPath As String
    Dim WLLng_Ret As Long
    
    WLStr_IniPath = IIf(Trim$(PPStr_Arquivo) <> "", PPStr_Arquivo, App.Path & "\KORSYS.ini")
    
    WLLng_Ret = WritePrivateProfileString(PPStr_Secao, PPStr_Chave, PPStr_Valor, WLStr_IniPath)
    GravarINI = (WLLng_Ret <> 0)
End Function


