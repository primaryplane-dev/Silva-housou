Option Explicit

'                                                          2016/11/30 Add 明細シートのcbo運送会社変更後を追加
'                                                          2017/05/03 Upd ３社固定
Public Sub Create運送会社リスト()
    Dim CN                  As Object
    Dim RS                  As Object
    Dim strSQL              As String
    Dim 行                  As Long

    'クリア
    '３社はよく使うのでリストの上に出したい
    With st01List.cbo運送会社
        .Clear
        .AddItem:   .List(0, 0) = "":       .List(0, 1) = ""
        .AddItem:   .List(1, 0) = "01":    .List(1, 1) = "名鉄運輸㈱"
        .AddItem:   .List(2, 0) = "02":    .List(2, 1) = "濃飛西濃運輸㈱"
        .AddItem:   .List(3, 0) = "14":    .List(3, 1) = "佐川急便"
    End With
    With st02Meisai.cbo運送会社変更後
        .Clear
        .AddItem:   .List(0, 0) = "":       .List(0, 1) = ""
        .AddItem:   .List(1, 0) = "01":    .List(1, 1) = "名鉄運輸㈱"
        .AddItem:   .List(2, 0) = "02":    .List(2, 1) = "濃飛西濃運輸㈱"
        .AddItem:   .List(3, 0) = "14":    .List(3, 1) = "佐川急便"
    End With
    行 = 3

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
        Do While (True)
            '読みとばし
            If RTrim(RS("YUCD")) = "" Then Exit Do
            If RTrim(RS("YUCD")) = "01" Then Exit Do
            If RTrim(RS("YUCD")) = "02" Then Exit Do
            If RTrim(RS("YUCD")) = "14" Then Exit Do

            'リストに加える
            行 = 行 + 1
            st01List.cbo運送会社.AddItem
            st01List.cbo運送会社.List(行, 0) = RTrim(RS("YUCD"))
            st01List.cbo運送会社.List(行, 1) = RTrim(RS("YUNM"))
            st02Meisai.cbo運送会社変更後.AddItem
            st02Meisai.cbo運送会社変更後.List(行, 0) = RTrim(RS("YUCD"))
            st02Meisai.cbo運送会社変更後.List(行, 1) = RTrim(RS("YUNM"))
            Exit Do
        Loop
        RS.MoveNext
    Loop

    'ＲＳクローズ、ＤＢ切断
    RS.Close:    Set RS = Nothing
    CN.Close:    Set CN = Nothing

End Sub


