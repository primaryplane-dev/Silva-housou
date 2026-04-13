Option Explicit

'「User.ini」の内容を取得する。
Public Sub ReadUserIni(ByVal iniFilePath As String)
    Dim strBuffer   As String
    Dim strKey      As String
    Dim strValue    As String

    Open iniFilePath For Input As #1
    Do While Not EOF(1)
        Line Input #1, strBuffer
        If Len(Trim(strBuffer)) > 0 Then
            Call PS_Split(strBuffer, strKey, strValue)
            Select Case strKey
                Case "USR":   P_USER = strValue
                Case "KGK":   P_権限 = strValue
            End Select
        End If
    Loop
    Close #1
       
End Sub

'文字列分割
Public Sub PS_Split(ByVal i_strString As String, ByRef o_strKey As String, ByRef o_strValue As String)
    Dim intStart    As Integer
    
    intStart = InStr(1, i_strString, "=")
    o_strKey = Left(i_strString, intStart - 1)
    o_strValue = Mid(i_strString, intStart + 1)
End Sub

