Attribute VB_Name = "modINI"
Private Declare Function GetIniInfo% Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal KeyDefault$, ByVal Returnstring$, ByVal NumBytes As Integer, ByVal FileName$)
Private Declare Function WriteIniInfo Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetIniData(strAppName$, strKeyName$, strDefault$, strIniFileName$) As String
    Dim INIBuffer As String
    Dim INIBufferSize As Long
    Dim BytesReturned As Integer

    INIBuffer = Space$(128)
    INIBufferSize = Len(INIBuffer)
    BytesReturned = GetIniInfo(strAppName$, strKeyName$, strDefault$, INIBuffer, INIBufferSize, strIniFileName$)
    If BytesReturned = 0 Then
        GetIniData = ""
    Else
        GetIniData = Mid$(INIBuffer, 1, BytesReturned)
    End If
End Function

Public Function WriteIniData(strAppName$, strKeyName$, strData$, strIniFileName$) As Boolean

    Dim BytesReturned As Integer

    INIBuffer = Space$(128)
    INIBufferSize = Len(INIBuffer)
    BytesReturned = WriteIniInfo(strAppName$, strKeyName$, strData$, strIniFileName$)
End Function

