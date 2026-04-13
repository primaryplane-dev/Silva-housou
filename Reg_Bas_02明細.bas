Option Explicit

Public Const 明細_行頭 = 11
Public 明細_最終行          As Long

Public Sub 明細クリア()

    st02Meisai.Select

    'クリア
    st02Meisai.Cells.Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Selection.Borders.LineStyle = xlLineStyleNone
    st02Meisai.cbo運送会社変更後.ListIndex = -1

    '列みだし
    Cells(1, 1) = "出荷明細表"
    Cells(3, 2) = "出荷日":     Cells(3, 7) = "温度帯"
    Cells(4, 2) = "納品日":     Cells(4, 7) = "伝票No."
    Cells(5, 2) = "出荷先":     Cells(5, 7) = "運送会社"
    Cells(6, 2) = "納品先":     Cells(6, 7) = " 運送会社 変更後"
    Cells(8, 2) = "メモ"
    '列みだし(明細)
    Cells(10, 2) = "行番号"
    Cells(10, 3) = "伝票区分"
    Cells(10, 4) = "品番"
    Cells(10, 5) = "品名"
    Cells(10, 6) = "入数"
    Cells(10, 7) = "単位"
    Cells(10, 8) = "注文数"
    Cells(10, 9) = "引当在庫(賞味期限)"
    Cells(10, 10) = "チェック"
    Cells(10, 11) = "車両積荷前衛生点検"
    Cells(10, 12) = "逸脱事項"
    Range(Cells(3, 2), Cells(8, 2)).Interior.Color = RGB(255, 255, 153)     '薄黄
    Range(Cells(3, 7), Cells(5, 7)).Interior.Color = RGB(255, 255, 153)     '薄黄
    Range(Cells(6, 7), Cells(6, 8)).Interior.Color = RGB(255, 255, 153)     '薄黄
    Range(Cells(3, 2), Cells(5, 2)).Borders.LineStyle = xlContinuous
    Range(Cells(3, 7), Cells(5, 7)).Borders.LineStyle = xlContinuous
    Range(Cells(6, 2), Cells(7, 2)).BorderAround Weight:=xlThin
    Range(Cells(8, 2), Cells(8, 2)).BorderAround Weight:=xlThin
    Range(Cells(6, 7), Cells(6, 8)).BorderAround Weight:=xlThin
    '明細部
    Range(Cells(10, 2), Cells(10, 8)).Interior.Color = RGB(255, 255, 153)   '薄黄
    Range(Cells(10, 9), Cells(10, 12)).Interior.Color = RGB(255, 204, 153)  '薄橙
    Range(Cells(10, 2), Cells(10, 12)).Borders.LineStyle = xlContinuous
    
    'テスト時は警告表示する(タイトル行をオレンジ色に)
    If P_LIB = "LIBSMF17T" Then Range(Cells(1, 1), Cells(1, 14)).Interior.Color = RGB(255, 100, 0)
    
    Cells(1, 7).Select
    
End Sub

Public Sub 明細表示()
    Dim CN                  As Object
    Dim RS                  As Object
    Dim strSQL              As String
    Dim 行                  As Long
    Dim data行              As Long
    Dim KEY                 As String
    Dim KEY_Z               As String
    Dim WKメモ              As String
    Dim WK今回引当          As String
    Dim WK期限切れ          As String
    Dim WK割当数計          As Long

    st02Meisai.Select
    Call 明細クリア
    
    'ＤＢ接続
    Set CN = CreateObject("ADODB.Connection")
    Set RS = CreateObject("ADODB.Recordset")
    CN.CursorLocation = adUseClient
    CN.Open P_接続文字列

    '■ヘッダデータ抽出　LIBWMF17.WNPP21B3:受注・納品プール
    '  (伝票情報)        LIBWMF.WMSP01    :名称マスタ
    '                  　LIBBMF.BAEP01  　:住所マスタ
    '                    LIBWMF17.WTMP01  :特約店マスタ(出荷先)
    '                  　LIBWMF17.WTEP01  :特約店枝番管理マスタ
    行 = 明細_行頭 - 1
    strSQL = ""
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "   JP.*  "
    strSQL = strSQL & "  ,MSRYM AS ONDO "
    strSQL = strSQL & "  ,AEIKM "
    strSQL = strSQL & "  ,TMKTM "
    strSQL = strSQL & "  ,TETEL "
    strSQL = strSQL & "  ,TEME1 "
    strSQL = strSQL & "  ,TEME2 "
    strSQL = strSQL & "  ,YUCA "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & "  ( SELECT * "
    strSQL = strSQL & "    FROM  LIBWMF17.WNPP21B3 "
    strSQL = strSQL & "    WHERE JPSNO='" & P_専用伝票NO & "' "
    strSQL = strSQL & "    FETCH FIRST 1 ROWS ONLY "
    strSQL = strSQL & "  ) AS JP "
    strSQL = strSQL & " LEFT JOIN LIBWMF.WMSP01    ON MSMSC=JPODK AND MSMSK='09' "
    strSQL = strSQL & " LEFT JOIN LIBBMF.BAEP01    ON AEANO=CONCAT(SUBSTR(JPETC,1,2), '000') "
    strSQL = strSQL & " LEFT JOIN LIBWMF17.WTMP01  ON TMTNO=JPTNO "
    strSQL = strSQL & " LEFT JOIN LIBWMF17.WTEP01  ON TETNO=JPTNO AND TEENO=JPSWK AND TECD1=JPHC4 "
    strSQL = strSQL & " LEFT JOIN "                                                                     '2016/11/30 Add
    strSQL = strSQL & "  ( SELECT ZSSNO,MAX(ZSYUCA) AS YUCA "                                           '全行同じ前提
    strSQL = strSQL & "      FROM " & P_LIB & ".SZSP01 "
    strSQL = strSQL & "     WHERE ZSDLT='' "
    strSQL = strSQL & "       AND ZSSNO='" & P_専用伝票NO & "' "
    strSQL = strSQL & "     GROUP BY ZSSNO "
    strSQL = strSQL & "  ) AS ZS ON ZS.ZSSNO = JPSNO "
    strSQL = strSQL & " ORDER BY JPSNO,JPSGY "
    Debug.Print strSQL
    RS.Open strSQL, CN, adOpenStatic, adLockReadOnly
    Application.StatusBar = "データをシートに設定しています．．．"
    Do While Not RS.EOF
        Cells(3, 3) = RS("JPPNS") & RS("JPPNE") & "/" & RS("JPPTU") & "/" & RS("JPPHI") '出荷日
        Cells(3, 8) = RS("ONDO")                                                        '温度帯区分
        Cells(4, 3) = RS("JPNNS") & RS("JPNNE") & "/" & RS("JPNTU") & "/" & RS("JPNHI") '納品日
        Cells(4, 4) = RS("JPCNO") & "-" & RS("JPDNO")                                   'コースNO + 配送NO
        Cells(4, 8) = "'" & RTrim(RS("JPSNO"))                                          '専用伝票NO         2018/05/09　ゼロサプレス対応
        Cells(5, 3) = RS("JPTNO")                                                       '店舗NO
        Cells(5, 4) = RS("JPHC4") & " " & RS("JPSWK")                                   '汎用CD4 + 仕分区分
        Cells(5, 5) = RTrim(RS("TMKTM"))                                                '出荷先名
        Cells(5, 9) = P_運送会社NM
        Cells(5, 8) = "【" & RS("AEIKM") & "】"                                         '県名
        Cells(6, 3) = RTrim(RS("TETEL"))                                                '納品先TEL
        Cells(6, 5) = RTrim(RS("TEME1"))                                                '納品先名
        Cells(7, 5) = RTrim(RS("TEME2"))                                                '納品先住所
        Cells(8, 3) = RTrim(RS("JPMEM"))                                                'メモ
        If Trim(NVL(RS("YUCA"))) <> "" Then
            st02Meisai.cbo運送会社変更後.Value = NVL(RS("YUCA"))
        End If
        Exit Do
    Loop
    'ＲＳクローズ、ＤＢ切断
    RS.Close:    Set RS = Nothing
    CN.Close:    Set CN = Nothing

    '■在庫引当シートを転記する(KEYは行NO)
    行 = 明細_行頭 - 1
    '変数クリア
    WKメモ = "":    WK割当数計 = 0:    WK今回引当 = "": WK期限切れ = ""
    KEY_Z = st02Hikiate.Cells(引当_行頭, 3)
    For data行 = 引当_行頭 To 引当_最終行 + 1
        KEY = st02Hikiate.Cells(data行, 3)
        If KEY <> KEY_Z Then
            'ブレイクしたら前の行の情報を書く
            行 = 行 + 1
            Cells(行, 2) = st02Hikiate.Cells(data行 - 1, 3)     '伝票No.行番号
            Cells(行, 3) = st02Hikiate.Cells(data行 - 1, 4)     '伝票区分
            Cells(行, 4) = st02Hikiate.Cells(data行 - 1, 5)     '品番
            Cells(行, 5) = st02Hikiate.Cells(data行 - 1, 6)     '品名
            Cells(行, 6) = st02Hikiate.Cells(data行 - 1, 7)     '入数
            Cells(行, 7) = st02Hikiate.Cells(data行 - 1, 9)     '単位名
            Cells(行, 8) = st02Hikiate.Cells(data行 - 1, 10)    '注文数
            Cells(行, 9) = WKメモ                               '引当在庫(賞味期限)
            If WK割当数計 >= st02Hikiate.Cells(data行 - 1, 10) Then
                If WK今回引当 = "あり" Then
                    Cells(行, 10) = "未処理"                    'チェック
                Else
                    Cells(行, 10) = "確定"
                End If
            End If
            If WK期限切れ = "あり" Then
                Cells(行, 11) = "期限ぎれ在庫あり"              '欄外コメント
            End If
            '変数クリア
            WKメモ = "":        WK割当数計 = 0
            WK期限切れ = "":    WK今回引当 = ""
            KEY_Z = KEY
        End If
    
        '出荷・出荷候補数をまとめる
        If Val(st02Hikiate.Cells(data行, 14)) <> 0 Then
            WK割当数計 = WK割当数計 + st02Hikiate.Cells(data行, 14)
            If WKメモ <> "" Then WKメモ = WKメモ & vbCrLf
                                                                            '2017/05/08 Upd No.62 バッチNoを削除
           'WKメモ = WKメモ & Right(String(6, " ") & st02Hikiate.Cells(data行, 14), 6) _
           '                & " (" _
           '                & Format(Get賞味期限fromロット(st02Hikiate.Cells(data行, 16)), "yyyy/mm/dd") _
           '                & "-" & Getバッチ数fromロット(st02Hikiate.Cells(data行, 16)) _
           '                & ") " _
           '                & st02Hikiate.Cells(data行, 15)
            WKメモ = WKメモ & Right(String(6, " ") & st02Hikiate.Cells(data行, 14), 6) _
                            & " (" _
                            & Format(Get賞味期限fromロット(st02Hikiate.Cells(data行, 16)), "yyyy/mm/dd") _
                            & ") " _
                            & st02Hikiate.Cells(data行, 15)
            If st02Hikiate.Cells(data行, 15) = "*" Then WK今回引当 = "あり"
        End If
        If st02Hikiate.Cells(data行, 15) = "x" Then WK期限切れ = "あり"
    Next
    明細_最終行 = 行

    '見ためを整える
    Range(Cells(明細_行頭, 2), Cells(明細_最終行, 12)).Rows.AutoFit
    Range(Cells(明細_行頭, 2), Cells(明細_最終行, 12)).Borders.LineStyle = xlContinuous

    Cells(1, 7).Select
    Application.StatusBar = False

End Sub



