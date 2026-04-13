Option Explicit

Public リスト_最終行        As Long
Public Const リスト_開始行 = 6
Private Const 開始列 = 2
Private Const 終了列 = 4

Public Sub 一覧クリア()
'Private Sub 一覧クリア()

    st01List.Activate
    Cells.Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Selection.Borders.LineStyle = xlLineStyleNone

    '列みだし
    Cells(1, 1) = "出荷トレース"
    Cells(5, 2) = "品番"
    Cells(5, 3) = "品名"
    Cells(5, 4) = "賞味期限"

    'ライブラリを判定する
    If P_LIB = "LIBSMF17T" Then Range(Cells(1, 1), Cells(1, 10)).Interior.Color = RGB(255, 100, 0)

End Sub

Public Sub 一覧表示()
    Dim CN                  As Object
    Dim RS                  As Object
    Dim strSQL              As String
    Dim 行                  As Long
    Dim i                   As Integer
    Dim strWK               As String
    
    st01List.Select
    Call 一覧クリア
    
    '検索条件
    If P_賞味期限日 > 0 Then Cells(2, 3).Value = Format(P_賞味期限日, "yyyy/mm/dd")
    If UBound(P_製品) > 0 Then
        For i = 1 To UBound(P_製品)
            If i > 1 Then strWK = strWK & ", "
            strWK = strWK & P_製品(i).CD
        Next
    End If
    Cells(3, 3).Value = strWK

    行 = リスト_開始行 - 1

    'ＤＢ接続
    Application.StatusBar = "データをＤＢから抽出しています．．．"
    Set CN = CreateObject("ADODB.Connection")
    Set RS = CreateObject("ADODB.Recordset")
    CN.CursorLocation = adUseClient
    CN.Open P_接続文字列
    
    '■データ抽出　　　　LIBSMF17.SRHP01  :製品マスタ
    '                    LIBSMF17.SSZP01  :在庫データ
    strSQL = ""
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "   RHHNO AS SNO  "
    strSQL = strSQL & "  ,SZHNO AS HNO  "
    strSQL = strSQL & "  ,RHHNM AS HNM  "
    strSQL = strSQL & "  ,SZLOT AS LOT  "
    strSQL = strSQL & "  ,SZLMT AS LMT"
    strSQL = strSQL & " FROM " & P_LIB & ".SRHP01 "
    strSQL = strSQL & " INNER JOIN " & P_LIB & ".SSZP01 ON SZSNO = RHHNO AND SZDLT='' "
    If P_賞味期限日 > 0 Then strSQL = strSQL & "       AND SZLMT = " & Format(P_賞味期限日, "yyyymmdd")
    strSQL = strSQL & " WHERE RHDLT='' "
    If UBound(P_製品) > 0 Then
        strSQL = strSQL & " AND ( "
        For i = 1 To UBound(P_製品)
            If i > 1 Then strSQL = strSQL & " OR "
            strSQL = strSQL & " SZSNO='" & P_製品(i).CD & "' "
        Next
        strSQL = strSQL & " ) "
    End If
    strSQL = strSQL & " ORDER BY RHHNO,LOT "
    Debug.Print strSQL
    RS.Open strSQL, CN, adOpenStatic, adLockReadOnly
    Application.StatusBar = "データをシートに設定しています．．．"
    Do While Not RS.EOF
        行 = 行 + 1
        Cells(行, 2) = RS("SNO")                                '品番
        Cells(行, 3) = RS("HNM")                                '品名
        Cells(行, 4) = 日付変換(RS("LMT"))                      '賞味期限
        RS.MoveNext
    Loop
    リスト_最終行 = 行
    Cells(5, 1).Value = リスト_最終行

    'ＲＳクローズ、ＤＢ切断
    RS.Close:    Set RS = Nothing
    CN.Close:    Set CN = Nothing

    '見ためを整える
    Range(Cells(リスト_開始行 - 1, 開始列), Cells(リスト_開始行 - 1, 終了列)).Interior.Color = RGB(255, 255, 204)
    Range(Cells(リスト_開始行 - 1, 開始列), Cells(リスト_最終行, 終了列)).Borders.LineStyle = xlContinuous
    
    Cells(1, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Application.StatusBar = False

End Sub

