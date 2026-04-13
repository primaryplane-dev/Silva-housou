Option Explicit

'数量は、少数点がない前提で作っています(簡易的にVal関数を使用)

Public Const 引当_行頭 = 4
Public 引当_最終行          As Long

'在庫引当用ワーク
Public Type 引当Record
    '注文
    伝票NO      As String
    行NO        As String
    伝票区分    As String
    販売品番    As String
    販売品名    As String
    入数        As String
    単位        As String
    単位名      As String
    注文数      As Long
    '出荷/在庫
    販売品番2   As String
    生産品番    As String
    在庫数      As Long
    出荷数      As Long
    区分        As String
    ロット      As String
    賞味期限    As Date
    バッチNO    As String
    出庫期限    As Date
End Type

Public Sub 在庫引当クリア()

    st02Hikiate.Activate

    'クリア
    st02Hikiate.Cells.Select
    Selection.ClearContents
    Selection.Font.ColorIndex = xlAutomatic
    Selection.Interior.ColorIndex = xlNone
    Selection.Borders.LineStyle = xlLineStyleNone

    '列みだし
    Cells(1, 1) = "在庫引当ワーク"
    Cells(2, 2) = "注文"
    Cells(3, 2) = "伝票No."
    Cells(3, 3) = "行番号"
    Cells(3, 4) = "伝票区分"
    Cells(3, 5) = "販売品番"
    Cells(3, 6) = "販売品名"
    Cells(3, 7) = "入数"
    Cells(3, 8) = "単位"
    Cells(3, 9) = "単位名"
    Cells(3, 10) = "注文数"
    Cells(2, 11) = "出荷/在庫"
    Cells(3, 11) = "販売品番"
    Cells(3, 12) = "生産品番"
    Cells(3, 13) = "在庫数"
    Cells(3, 14) = "出荷数"
    Cells(3, 15) = "仮"
    Cells(3, 16) = "ロットNo."      '賞味期限＋バッチNo.
    Cells(3, 17) = "出庫期限"
    Range(Cells(2, 2), Cells(3, 10)).Interior.Color = RGB(255, 255, 153)     '黄
    Range(Cells(2, 11), Cells(3, 17)).Interior.Color = RGB(255, 153, 204)    '桃
    Range(Cells(2, 2), Cells(3, 17)).Borders.LineStyle = xlContinuous
    Cells(1, 7).Select
    
End Sub

Public Sub Create在庫引当ワーク()
    Dim 行                  As Long
    Dim data行              As Long
    Dim KEY                 As String
    Dim KEY_Z               As String
    Dim 引当Rec()           As 引当Record

    st02Hikiate.Activate
    Call 在庫引当クリア

    'データ抽出
    Call データ抽出_出荷在庫(引当Rec, "")
    
    行 = 引当_行頭 - 1
    For data行 = 1 To UBound(引当Rec)
        行 = 行 + 1
        With 引当Rec(data行)
            '注文
            Cells(行, 2) = .伝票NO
            Cells(行, 3) = .行NO                                        '伝票No.の行番号
            Cells(行, 4) = .伝票区分
            Cells(行, 5) = .販売品番
            Cells(行, 6) = .販売品名
            Cells(行, 7) = .入数
            Cells(行, 8) = .単位
            Cells(行, 9) = .単位名
            Cells(行, 10) = .注文数
            '出荷/在庫
            Cells(行, 11) = .販売品番2
            Cells(行, 12) = .生産品番
            If .在庫数 <> 0 Then Cells(行, 13) = .在庫数
            If .出荷数 <> 0 Then Cells(行, 14) = .出荷数
            Cells(行, 15) = .区分
            Cells(行, 16) = .ロット
            If .出庫期限 > 0 Then Cells(行, 17) = .出庫期限
        End With
    Next
    引当_最終行 = 行

    '見ためを整える
    Range(Cells(引当_行頭, 2), Cells(引当_最終行, 17)).Borders.LineStyle = xlContinuous
    For 行 = 引当_行頭 To 引当_最終行
        KEY = Cells(行, 2) & Cells(行, 3) & Cells(行, 4) & Cells(行, 5) & Cells(行, 6) & Cells(行, 7)
        If KEY = KEY_Z Then
            Range(Cells(行, 2), Cells(行, 10)).Font.Color = RGB(192, 192, 192)
        Else
            KEY_Z = KEY
        End If
    Next

    Cells(1, 7).Select
    Application.StatusBar = False

End Sub

Public Sub データ抽出_出荷在庫(ByRef 引当Rec() As 引当Record, ByVal i_行NO As String)
    Dim CN                  As Object
    Dim RS                  As Object
    Dim strSQL              As String
    Dim 行                  As Long
    
    'クリア
    ReDim 引当Rec(0)

    'ＤＢ接続
    Set CN = CreateObject("ADODB.Connection")
    Set RS = CreateObject("ADODB.Recordset")
    CN.CursorLocation = adUseClient
    CN.Open P_接続文字列
    
    '■データ抽出　　　　LIBWMF17.WNPP21B3:注文データ
    '                    LIBSMF17.SZSP01  :出荷データ
    '                    LIBSMF17.SSZP01  :在庫データ
    '出荷
    strSQL = ""
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "    JPSNO   AS DPNO "    '1
    strSQL = strSQL & "   ,JPSGY   AS GYO  "    '2
    strSQL = strSQL & "   ,JPDPK   AS DPK  "    '3
    strSQL = strSQL & "   ,JPHNO   AS HNO  "    '4
    strSQL = strSQL & "   ,JPHNM   AS HNM  "    '5
    strSQL = strSQL & "   ,JPIRS   AS IRS  "    '6
    strSQL = strSQL & "   ,JPTNI   AS TNI  "    '7
    strSQL = strSQL & "   ,JPTNN   AS TNN  "    '8
    strSQL = strSQL & "   ,JPKSU   AS KSU  "    '9
    strSQL = strSQL & "   ,JPNNS || JPNNE || right ('00' || JPNTU, 2) || right ('00' || JPNHI, 2) AS NHI "
    strSQL = strSQL & "   ,ZSHNO   AS HNO2 "    '11
    strSQL = strSQL & "   ,ZSHNO   AS SNO  "    '12     '※原則 生産品番＝販売品番
    strSQL = strSQL & "   ,ZSSRY   AS SRY  "    '13
    strSQL = strSQL & "   ,ZSLOT   AS LOT  "    '14
    strSQL = strSQL & "   ,'出荷'  AS KBN  "    '15
    strSQL = strSQL & "   ,'1'     AS KBN2 "    '16 KBNのSORT用
    strSQL = strSQL & "   ,0       AS SLD  "    '17
    strSQL = strSQL & "   ,0       AS SLD2  "   '18
    ' 2018/05/16  伝票番号+行番号で受注データが２件以上あるケース。条件に日付も入れておく
    'strSQL = strSQL & " FROM      LIBWMF17.WNPP21B3 "
    strSQL = strSQL & " FROM "
    ' 2021/10/15  日付+伝票番号+行番号で受注データが２件以上あるケース。条件に品番を入れゼロは省く（打消しの伝票）
    strSQL = strSQL & "    (SELECT"
    strSQL = strSQL & "             *"
    strSQL = strSQL & "     From"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "              JPSNO"
    strSQL = strSQL & "             ,JPSGY"
    strSQL = strSQL & "             ,MAX(JPDPK) AS JPDPK"
    strSQL = strSQL & "             ,MAX(JPHNO) AS JPHNO"
    strSQL = strSQL & "             ,MAX(JPHNM) AS JPHNM"
    strSQL = strSQL & "             ,MAX(JPIRS) AS JPIRS"
    strSQL = strSQL & "             ,MAX(JPTNI) AS JPTNI"
    strSQL = strSQL & "             ,MAX(JPTNN) AS JPTNN"
    strSQL = strSQL & "             ,SUM(JPKSU) AS JPKSU"
    strSQL = strSQL & "             ,JPNNS"
    strSQL = strSQL & "             ,JPNNE"
    strSQL = strSQL & "             ,JPNTU"
    strSQL = strSQL & "             ,JPNHI"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             LIBWMF17.WNPP21B3"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             JPSNO='" & P_専用伝票NO & "' "
    'strSQL = strSQL & "         GROUP BY JPNNS,JPNNE,JPNTU,JPNHI,JPSNO,JPSGY"
    strSQL = strSQL & "         GROUP BY JPNNS,JPNNE,JPNTU,JPNHI,JPSNO,JPSGY,JPHNO"
    strSQL = strSQL & "      ) AS JP2"
    strSQL = strSQL & "     WHERE JPKSU > 0"
    strSQL = strSQL & "    ) AS JP"
    strSQL = strSQL & " LEFT JOIN "
    strSQL = strSQL & "  ( SELECT * "
    strSQL = strSQL & "      FROM " & P_LIB & ".SZSP01 "
    strSQL = strSQL & "     WHERE ZSDLT='' "
    strSQL = strSQL & "       AND ZSSNO='" & P_専用伝票NO & "' "
    If Val(i_行NO) > 0 Then strSQL = strSQL & " AND ZSSGY=" & i_行NO
    strSQL = strSQL & "  ) AS ZS ON ZS.ZSSNO = JP.JPSNO AND ZS.ZSSGY = JP.JPSGY "
    strSQL = strSQL & " WHERE JPSNO='" & P_専用伝票NO & "' "
    If Val(i_行NO) > 0 Then strSQL = strSQL & " AND JPSGY=" & i_行NO
    '在庫
    strSQL = strSQL & " UNION ALL "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "    JPSNO   AS DPNO "
    strSQL = strSQL & "   ,JPSGY   AS GYO  "
    strSQL = strSQL & "   ,JPDPK   AS DPK  "
    strSQL = strSQL & "   ,JPHNO   AS HNO  "
    strSQL = strSQL & "   ,JPHNM   AS HNM  "
    strSQL = strSQL & "   ,JPIRS   AS IRS  "
    strSQL = strSQL & "   ,JPTNI   AS TNI  "
    strSQL = strSQL & "   ,JPTNN   AS TNN  "
    strSQL = strSQL & "   ,JPKSU   AS KSU  "
    strSQL = strSQL & "   ,JPNNS || JPNNE || right ('00' || JPNTU, 2) || right ('00' || JPNHI, 2) AS NHI "
    strSQL = strSQL & "   ,SZHNO   AS HNO2 "
    strSQL = strSQL & "   ,SZSNO   AS SNO  "
    strSQL = strSQL & "   ,SZSRY   AS SRY  "
    strSQL = strSQL & "   ,SZLOT   AS LOT  "
    strSQL = strSQL & "   ,'在庫'  AS KBN  "
    strSQL = strSQL & "   ,'2'       AS KBN2 "    '16 KBNのSORT用
    strSQL = strSQL & "   ,SZSLD   AS SLD  "
    strSQL = strSQL & "   ,SZSLD2  AS SLD2 "                                    '2017/04/03 Add
    ' 2018/05/16  伝票番号+行番号で受注データが２件以上あるケース。条件に日付も入れておく
    'strSQL = strSQL & " FROM      LIBWMF17.WNPP21B3 "
    strSQL = strSQL & " FROM "
    ' 2021/10/15  日付+伝票番号+行番号で受注データが２件以上あるケース。条件に品番を入れゼロは省く（打消しの伝票）
    strSQL = strSQL & "    (SELECT"
    strSQL = strSQL & "             *"
    strSQL = strSQL & "     From"
    strSQL = strSQL & "         (SELECT"
    strSQL = strSQL & "              JPSNO"
    strSQL = strSQL & "             ,JPSGY"
    strSQL = strSQL & "             ,MAX(JPDPK) AS JPDPK"
    strSQL = strSQL & "             ,MAX(JPHNO) AS JPHNO"
    strSQL = strSQL & "             ,MAX(JPHNM) AS JPHNM"
    strSQL = strSQL & "             ,MAX(JPIRS) AS JPIRS"
    strSQL = strSQL & "             ,MAX(JPTNI) AS JPTNI"
    strSQL = strSQL & "             ,MAX(JPTNN) AS JPTNN"
    strSQL = strSQL & "             ,SUM(JPKSU) AS JPKSU"
    strSQL = strSQL & "             ,JPNNS"
    strSQL = strSQL & "             ,JPNNE"
    strSQL = strSQL & "             ,JPNTU"
    strSQL = strSQL & "             ,JPNHI"
    strSQL = strSQL & "         FROM"
    strSQL = strSQL & "             LIBWMF17.WNPP21B3"
    strSQL = strSQL & "         WHERE"
    strSQL = strSQL & "             JPSNO='" & P_専用伝票NO & "' "
    'strSQL = strSQL & "         GROUP BY JPNNS,JPNNE,JPNTU,JPNHI,JPSNO,JPSGY"
    strSQL = strSQL & "         GROUP BY JPNNS,JPNNE,JPNTU,JPNHI,JPSNO,JPSGY,JPHNO"
    strSQL = strSQL & "      ) AS JP2"
    strSQL = strSQL & "     WHERE JPKSU > 0"
    strSQL = strSQL & "    ) AS JP"
    strSQL = strSQL & " LEFT JOIN "
    strSQL = strSQL & "  ( SELECT * "
    strSQL = strSQL & "      FROM " & P_LIB & ".SSZP01 "
    strSQL = strSQL & "     WHERE SZDLT='' "
    strSQL = strSQL & "       AND SZSRY>0 "
    strSQL = strSQL & "  ) AS SZ ON SZ.SZSNO = JP.JPHNO "
    strSQL = strSQL & " WHERE JPSNO='" & P_専用伝票NO & "' "
    If Val(i_行NO) > 0 Then strSQL = strSQL & " AND JPSGY=" & i_行NO
    strSQL = strSQL & " ORDER BY 1,2,16,14 "                                    'SNO,SGY,KBN2,LOT
    Debug.Print strSQL
    RS.Open strSQL, CN, adOpenStatic, adLockReadOnly
    Do While Not RS.EOF
        If IsNull(RS("LOT")) And IsNull(RS("SRY")) And RS("KBN") = "在庫" Then
        Else
            行 = 行 + 1:    ReDim Preserve 引当Rec(行)
            With 引当Rec(行)
                '注文
                .伝票NO = "'" & RS("DPNO")                                      '2018/05/09　ゼロサプレス対応
                .行NO = RS("GYO")
                .伝票区分 = RS("DPK")
                .販売品番 = RS("HNO")
                .販売品名 = RTrim(RS("HNM"))
                .入数 = RS("IRS")
                .単位 = RS("TNI")
                If Trim(RS("TNN")) <> "" Then
                    .単位名 = RS("TNI") & "(" & Trim(RS("TNN")) & ")"
                End If
                .注文数 = RS("KSU")
                '出荷/在庫
                .販売品番2 = NVL(RS("HNO2"))
                .生産品番 = NVL(RS("SNO"))
                .出庫期限 = 日付変換(RS("SLD"))
                If P_出荷期限KB = "2" Then .出庫期限 = 日付変換(RS("SLD2")) '2017/04/03 Upd 出荷期限パターンでどちらかを使う
                If RS("KBN") = "出荷" Then
                    .在庫数 = 0
                    .出荷数 = NZ(RS("SRY"))
                    .区分 = "'―"
                    If .出荷数 <> 0 Then
                        .区分 = "確"
                        .ロット = NVL(RS("LOT"))
                        .賞味期限 = Get賞味期限fromロット(.ロット)
                        .バッチNO = Getバッチ数fromロット(.ロット)
                    End If
                End If
                If RS("KBN") = "在庫" Then
                    .在庫数 = RS("SRY")
                    .出荷数 = 0
                    If .在庫数 > 0 Then                             '2016/10/27 榊原 出荷期限が入ってない場合の処理を追加
                        Select Case True
                        Case .出庫期限 = 0
                            .区分 = "+"                            '出庫期限が未設定のときは出荷可能とみなす
                        Case .出庫期限 < 日付変換(RS("NHI")):
                            .区分 = "x"                            '出荷期限を過ぎている在庫
                        Case Else
                            .区分 = "+"                            '出荷可能な在庫
                        End Select
                        .ロット = NVL(RS("LOT"))
                        .賞味期限 = Get賞味期限fromロット(.ロット)
                        .バッチNO = Getバッチ数fromロット(.ロット)
                    End If


                End If
            End With
        End If
        RS.MoveNext
    Loop
    
    '■在庫引当(自動)　注文数に在庫をあてはめていく
    Dim KEY         As String
    Dim KEY_Z       As String
    Dim WK注文数    As Long
    Dim WK出荷数    As Long
    Dim WK必要数    As Long
    Dim WK割当数    As Long
    '初期値
    KEY_Z = KEY
    WK注文数 = Val(Cells(1, 10))
    '初期値
    For 行 = 1 To UBound(引当Rec)
        With 引当Rec(行)
            KEY = .伝票NO & .行NO & .伝票区分 & .販売品番
            If KEY <> KEY_Z Then
                KEY_Z = KEY
                WK注文数 = .注文数
                WK出荷数 = 0
            End If
            '出荷数を集計する
            If .区分 = "確" Then
                WK出荷数 = WK出荷数 + .出荷数
            End If
            '在庫を引き当てる
            If .区分 = "+" Then
                WK必要数 = WK注文数 - WK出荷数
                If WK必要数 > 0 Then
                    WK割当数 = IIf(WK必要数 >= .在庫数, .在庫数, WK必要数)
                    WK出荷数 = WK出荷数 + WK割当数
                    .出荷数 = WK割当数
                    .区分 = "*"
                End If
            End If
        End With
    Next
       
    'ＲＳクローズ、ＤＢ切断
    RS.Close:    Set RS = Nothing
    CN.Close:    Set CN = Nothing

End Sub
