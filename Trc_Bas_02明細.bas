Option Explicit

Private Const 明細在庫_行頭 = 5
Private Const 明細出荷_行頭 = 8 '2017/05/01 Update

Private Sub 明細クリア()

    st02Meisai.Activate
    Cells.Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Selection.Borders.LineStyle = xlLineStyleNone

    '列みだし
    Cells(1, 1) = "出荷トレース"
    Cells(4, 1) = "在庫":        Cells(7, 1) = "出荷" '2017/05/01 Update
    Cells(4, 2) = "品番":        Cells(7, 2) = "出荷先" '2017/05/01 Update
    Cells(4, 3) = "品名":        Cells(7, 3) = "出荷先名" '2017/05/01 Update
    Cells(4, 4) = "賞味期限":    Cells(7, 4) = "納品先名" '2017/05/01 Update
    Cells(4, 5) = "出荷期限":    Cells(7, 5) = "運送会社" '2017/05/01 Update
    Cells(4, 6) = "出荷期限2":   Cells(7, 6) = "運送会社変更後" '2017/05/01 Update
    Cells(4, 7) = "ＪＡＮ":      Cells(7, 7) = "伝票No" '2017/05/01 Update
    Cells(4, 8) = "在庫調整日":  Cells(7, 8) = "納品日" '2017/05/01 Update
    Cells(4, 9) = "生産日":      Cells(7, 9) = "出荷日" '2017/05/01 Update
    Cells(4, 10) = "在庫数":     Cells(7, 10) = "出荷数" '2017/05/01 Update
    '--- 2026/04/15 新仕様対応: 11・12列目見出し追加（出荷部） ---
    Cells(7, 11) = "車両積荷前衛生点検(ZSSSTF)" '1:実施 0:未実施
    Cells(7, 12) = "逸脱事項(ZSIDJK)"           'フリー入力
    '---
    Range(Cells(明細在庫_行頭 - 1, 2), Cells(明細在庫_行頭 - 1, 10)).Interior.Color = RGB(255, 255, 204)
    Range(Cells(明細出荷_行頭 - 1, 2), Cells(明細出荷_行頭 - 1, 10)).Interior.Color = RGB(255, 255, 204)
    Range(Cells(明細在庫_行頭 - 1, 2), Cells(明細在庫_行頭 - 1, 10)).Borders.LineStyle = xlContinuous
    Range(Cells(明細出荷_行頭 - 1, 2), Cells(明細出荷_行頭 - 1, 10)).Borders.LineStyle = xlContinuous

End Sub

Public Sub 明細表示()
    Dim CN                  As Object
    Dim RS                  As Object
    Dim strSQL              As String
    Dim 行                  As Long
    
    st02Meisai.Select
    Call 明細クリア
        
    'ＤＢ接続
    Set CN = CreateObject("ADODB.Connection")
    Set RS = CreateObject("ADODB.Recordset")
    CN.CursorLocation = adUseClient
    CN.Open P_接続文字列
        
    行 = 明細在庫_行頭 - 1
    '■データ抽出①　　　LIBWMF17.SRHP01  :製品マスタ
    '                    LIBSMF17.SSZP01  :在庫データ
    strSQL = ""
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "   RHHNO AS SNO  "
    strSQL = strSQL & "  ,SZHNO AS HNO  "
    strSQL = strSQL & "  ,RHHNM AS HNM  "
    strSQL = strSQL & "  ,SZLOT AS LOT  "
    strSQL = strSQL & "  ,SZJAN AS JAN  "
    strSQL = strSQL & "  ,SZCDT AS CDT  "
    strSQL = strSQL & "  ,SZDAT AS DAT  "    '生産日
    strSQL = strSQL & "  ,SZLMT AS LMT  "    '賞味期限
    strSQL = strSQL & "  ,SZSLD AS SLD  "    '出荷期限
    strSQL = strSQL & "  ,SZSLD2 AS SLD2  "  '出荷期限2
    strSQL = strSQL & "  ,SZSRY AS SRY  "
    strSQL = strSQL & " FROM " & P_LIB & ".SRHP01 "
    strSQL = strSQL & " LEFT JOIN " & P_LIB & ".SSZP01 ON SZSNO = RHHNO AND SZDLT='' "
    strSQL = strSQL & " WHERE RHDLT='' "
    strSQL = strSQL & "       AND SZSNO='" & P_品番 & "' "
    strSQL = strSQL & "       AND SZLMT=" & Format(P_賞味期限, "yyyymmdd")
    strSQL = strSQL & " ORDER BY RHHNO,LOT "
    Debug.Print strSQL
    RS.Open strSQL, CN, adOpenStatic, adLockReadOnly
    Application.StatusBar = "データをシートに設定しています．．．" '2017/05/01
    Do While Not RS.EOF
        行 = 行 + 1
        Cells(行, 2) = RS("SNO")                                '品番
        Cells(行, 3) = RS("HNM")                                '品名
        Cells(行, 4) = 日付変換(RS("LMT"))                      '賞味期限
       'Cells(行, 5) = RS("LOT")                                'ロット         テスト用
        Cells(行, 5) = 日付変換(RS("SLD"))                      '出荷期限
        Cells(行, 6) = 日付変換(RS("SLD2"))                     '出荷期限2
        Cells(行, 7) = RS("JAN")                                'ＪＡＮ
        Cells(行, 8) = 日付変換(RS("CDT"))                      '在庫調整日
        Cells(行, 9) = 日付変換(RS("DAT"))                      '生産日
        Cells(行, 10) = RS("SRY")                               '現在庫数
        RS.MoveNext
    Loop
    RS.Close
    
    Range(Cells(明細在庫_行頭, 2), Cells(行, 10)).Borders.LineStyle = xlContinuous ' 2017/05/01 Add
    
    行 = 明細出荷_行頭 - 1
    '■データ抽出②　　　LIBSMF17.SZSP01  :出荷データ
    '                    LIBWMF17.WTMP01  :特約店マスタ(出荷先)
    '                  　LIBWMF17.WTEP01  :特約店枝番管理マスタ(納品先)
    '                  　LIBWMF.WSKP01    :倉庫マスタ(センター間用出荷先)
    '出荷
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "   ZSLOT   AS LOT  "
    strSQL = strSQL & "  ,ZSTNO   AS TNO  "
    strSQL = strSQL & "  ,ZSCPGM  AS PGM  "                                        '2018/04/12 Add
    strSQL = strSQL & "  ,ZSYUCD  AS YUCD "
    strSQL = strSQL & "  ,ZSYUCA  AS YUCA "
    strSQL = strSQL & "  ,ZSSNO   AS SNO  "
    strSQL = strSQL & "  ,ZSNDT   AS NDT  "
    strSQL = strSQL & "  ,ZSSDT   AS SDT  "
    strSQL = strSQL & "  ,ZSSRY   AS SRY  "
    strSQL = strSQL & "  ,TMKTM "
    strSQL = strSQL & "  ,TEME1 "
    strSQL = strSQL & "  ,SKSKM "                                                   '2018/04/12 Add
    strSQL = strSQL & " FROM      "
    strSQL = strSQL & "  ( SELECT "
    strSQL = strSQL & "      *  "
    strSQL = strSQL & "     FROM " & P_LIB & ".SZSP01 "
    strSQL = strSQL & "     WHERE ZSDLT='' "
    strSQL = strSQL & "       AND ZSHNO='" & P_品番 & "' "
    strSQL = strSQL & "       AND ZSLMT=" & Format(P_賞味期限, "yyyymmdd")
    strSQL = strSQL & "  ) AS ZS "
    strSQL = strSQL & " LEFT JOIN LIBWMF17.WNPP21B3 ON JPSNO = ZS.ZSSNO AND JPSGY = ZS.ZSSGY "
    strSQL = strSQL & " LEFT JOIN LIBWMF17.WTMP01   ON TMTNO=ZSTNO "
    strSQL = strSQL & " LEFT JOIN LIBWMF.WSKP01   ON RIGHT(TRIM('00'||CHAR(SKSKC)),3) = ZSTNO "                  '2018/04/12 Add
    strSQL = strSQL & " LEFT JOIN LIBWMF17.WTEP01   ON TETNO=ZSTNO AND TEENO=ZSSWK AND TECD1=ZSHC4 "
    Debug.Print strSQL
    RS.Open strSQL, CN, adOpenStatic, adLockReadOnly
    Do While Not RS.EOF
        行 = 行 + 1
        '一行おきに色をつける 2017/05/01
        If (行 Mod 2) = 0 Then Range(Cells(行, 2), Cells(行, 10)).Interior.Color = RGB(204, 255, 204)
        Cells(行, 2).Value = RS("TNO")                      '出荷先CD
        
        '出荷先名 (センター間は倉庫マスタの出荷先を編集する)                                                      '2018/04/12 Add
        If RS("PGM") = "ShukkaCXLS" Then
            Cells(行, 3).Value = RTrim(RS("SKSKM"))         '出荷先名（倉庫マスタ）
'            Cells(行, 4).Value = ""
        Else
            Cells(行, 3).Value = RTrim(RS("TMKTM"))         '出荷先名（特約店マスタ）
            Cells(行, 4).Value = RTrim(RS("TEME1"))         '納品先名
        End If
        Cells(行, 5).Value = Get運送会社NM(RS("YUCD"))      '運送会社
        Cells(行, 6).Value = Get運送会社NM(RS("YUCA"))      '運送会社変更後
        Cells(行, 7).Value = "'" & RTrim(RS("SNO"))         '専用伝票No.
        Cells(行, 8).Value = 日付変換(RS("NDT"))            '納品日
        Cells(行, 9).Value = 日付変換(RS("SDT"))            '出荷日
        Cells(行, 10).Value = RS("SRY")                     '出荷数
        '--- 2026/04/15 新仕様対応: 11・12列目（ZSSSTF, ZSIDJK）を転記 ---
        On Error Resume Next
        Cells(行, 11).Value = RS("ZSSSTF") '車両積荷前衛生点検（1:実施 0:未実施）
        Cells(行, 12).Value = RS("ZSIDJK") '逸脱事項（フリー入力）
        On Error GoTo 0
        '---
        RS.MoveNext
    Loop
    
    Range(Cells(明細出荷_行頭, 2), Cells(行, 10)).Borders.LineStyle = xlContinuous ' 2017/05/01 Add
    '--- 2026/04/15 新仕様対応: 11・12列目にも罫線を引く ---
    Range(Cells(明細出荷_行頭, 11), Cells(行, 12)).Borders.LineStyle = xlContinuous
    '---
    
    'ＲＳクローズ、ＤＢ切断
    RS.Close:    Set RS = Nothing
    CN.Close:    Set CN = Nothing

    Cells(1, 10).Select
    ActiveWindow.ScrollRow = 1 ' 2017/05/01 Add
    ActiveWindow.ScrollColumn = 1 ' 2017/05/01 Add
    Application.StatusBar = False ' 2017/05/01 Add
End Sub

