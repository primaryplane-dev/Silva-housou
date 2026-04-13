Option Explicit

Public Dic運送会社          As Object       '運送会社のCDと名称を結びつける

Public Sub Create運送会社リスト()
    Dim CN                  As Object
    Dim RS                  As Object
    Dim strSQL              As String
    Dim 行                  As Long

    Set Dic運送会社 = CreateObject("Scripting.Dictionary")
    
    'ＤＢ接続
    Set CN = CreateObject("ADODB.Connection")
    Set RS = CreateObject("ADODB.Recordset")
    CN.CursorLocation = adUseClient
    CN.Open P_接続文字列

    '■データ抽出　LIBWMF.WMSP01    :名称マスタ
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "   MSMSC  AS YUCD "
    strSQL = strSQL & "  ,MSMSM  AS YUNM "
    strSQL = strSQL & " FROM  LIBWMF.WMSP01 "
    strSQL = strSQL & " WHERE MSMSK='A6' "
    strSQL = strSQL & " ORDER BY 1 "
    Debug.Print strSQL
    RS.Open strSQL, CN, adOpenStatic, adLockReadOnly
    Do While Not RS.EOF
        If Not Dic運送会社.exists(RTrim(RS("YUCD"))) Then
            Dic運送会社.Add RTrim(RS("YUCD")), RTrim(RS("YUNM"))
        End If
        RS.MoveNext
    Loop

    'ＲＳクローズ、ＤＢ切断
    RS.Close:    Set RS = Nothing
    CN.Close:    Set CN = Nothing

End Sub

Public Function Get運送会社NM(ByVal i_CD As Variant) As String
    Get運送会社NM = ""

    If Dic運送会社 Is Nothing Then Call Create運送会社リスト

    If IsNull(i_CD) Then Exit Function
    
    Get運送会社NM = Dic運送会社.Item(RTrim(i_CD))
End Function



