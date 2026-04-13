Option Explicit

Public Const 出荷先_行頭 = 6
Public 出荷先_最終行        As Long

Private Const 開始列 = 2
Private Const 終了列 = 6

Public Sub 出荷先リストクリア()

    st01List.Select

    'クリア
    st01List.Cells.Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Selection.Borders.LineStyle = xlLineStyleNone

    '列みだし
    Cells(1, 1) = "包装出荷登録"
    Cells(2, 2) = "運送会社"
    Cells(3, 2) = "出荷日"
    Cells(出荷先_行頭 - 1, 2) = "出荷先"
    Cells(出荷先_行頭 - 1, 3) = "出荷先名"
    Cells(出荷先_行頭 - 1, 4) = "専用伝票No."
    Cells(出荷先_行頭 - 1, 5) = "受注数計"
    Cells(出荷先_行頭 - 1, 6) = "出荷数計"
    Cells(出荷先_行頭 - 1, 10) = "出荷期限パターン"
    Range(Cells(出荷先_行頭 - 1, 開始列), Cells(出荷先_行頭 - 1, 終了列)).Select
    Selection.Interior.Color = RGB(255, 255, 204)                                           '薄黄
    Selection.Borders.LineStyle = xlContinuous

    'テスト時は警告表示する(タイトル行をオレンジ色に)
    If P_LIB = "LIBSMF17T" Then Range(Cells(1, 1), Cells(1, 14)).Interior.Color = RGB(255, 100, 0)
    
    Cells(2, 7).Select
End Sub

Public Sub 出荷先リスト表示()
    Dim CN                  As Object
    Dim RS                  As Object
    Dim strSQL              As String
    Dim 行                  As Long

    st01List.Select
    行 = 出荷先_行頭 - 1
    Call 出荷先リストクリア

    '抽出条件(クリアされるため再設定)
    If P_出荷YMD > 0 Then Cells(3, 3) = Format(P_出荷YMD, "yyyy/mm/dd")

    'ＤＢ接続
    Application.StatusBar = "データをＤＢから抽出しています．．．"
    Set CN = CreateObject("ADODB.Connection")
    Set RS = CreateObject("ADODB.Recordset")
    CN.CursorLocation = adUseClient
    CN.Open P_接続文字列

    '■データ抽出　　　　LIBWMF17.WNPP21B3:受注・納品プール
    '                    LIBSMF17.SZSP01  :出荷データ
    '                    LIBWMF17.WTMP01  :特約店マスタ(出荷先)
    '                  　LIBWMF17.WTEP01  :特約店枝番管理マスタ
    '                    LIBSMF17.SRAP01  :出荷期限ルール適用マスタ
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "   JPTNO AS TNO "
    strSQL = strSQL & "  ,TMKTM AS TNM "
    strSQL = strSQL & "  ,TEME1 "
    strSQL = strSQL & "  ,PTN_N "
    strSQL = strSQL & "  ,PTN_T "
    strSQL = strSQL & "  ,JPSNO AS SNO "
    strSQL = strSQL & "  ,JUCHU_SU     "
    strSQL = strSQL & "  ,SHUKKA_SU    "
    strSQL = strSQL & " FROM  "
    strSQL = strSQL & "  ( SELECT "
    strSQL = strSQL & "      JPTNO  "
    strSQL = strSQL & "     ,JPSNO  "
    strSQL = strSQL & "     ,SUM(JPKSU) AS JUCHU_SU  "
    strSQL = strSQL & "     ,MAX(JPSWK) AS JPSWK  "
    strSQL = strSQL & "     ,MAX(JPHC4) AS JPHC4  "
    strSQL = strSQL & "    FROM LIBWMF17.WNPP21B3 "
    strSQL = strSQL & "    WHERE JPDLT='' "
    strSQL = strSQL & "      AND JPSYS='305'"                                       'ストア、BASEを除く
    strSQL = strSQL & "      AND CONCAT(JPHC9, JPHC0)='" & P_運送会社CD & "'"
    strSQL = strSQL & "      AND (JPPNE || right('00' || JPPTU,2) || right('00' || JPPHI,2)>=" & Format(P_出荷YMD, "yymmdd")
    strSQL = strSQL & "       AND JPPNE || right('00' || JPPTU,2) || right('00' || JPPHI,2)<=" & Format(P_出荷YMD, "yymmdd") & ")"
    strSQL = strSQL & "    GROUP BY JPTNO,JPSNO "
    strSQL = strSQL & "    HAVING SUM(JPKSU)>0 "
    strSQL = strSQL & "  ) AS X "
    strSQL = strSQL & " LEFT JOIN "
    strSQL = strSQL & "  ( SELECT "
    strSQL = strSQL & "      ZSSNO  "
    strSQL = strSQL & "     ,SUM(ZSSRY) AS SHUKKA_SU "
    strSQL = strSQL & "    FROM " & P_LIB & ".SZSP01 "
    strSQL = strSQL & "    WHERE ZSDLT='' "
    strSQL = strSQL & "      AND ZSSDT>=" & Format(P_出荷YMD, "yymmdd")
    strSQL = strSQL & "    GROUP BY ZSSNO "
    strSQL = strSQL & "  ) AS Y ON ZSSNO = X.JPSNO "
    strSQL = strSQL & " LEFT JOIN LIBWMF17.WTMP01 ON TMTNO = X.JPTNO "
    strSQL = strSQL & " LEFT JOIN  "
    strSQL = strSQL & "  ( SELECT  TETNO,TEENO,TECD1,TEME1   "
    strSQL = strSQL & "     FROM   LIBWMF17.WTEP01 "
    strSQL = strSQL & "     WHERE  TEDLT=''  "
    strSQL = strSQL & "  ) AS TE ON TETNO = X.JPTNO AND TEENO=X.JPSWK AND TECD1=X.JPHC4 "
    strSQL = strSQL & " LEFT JOIN  "
    strSQL = strSQL & "  ( SELECT  RATNO,RAENO,RACD1,RAPTN AS PTN_N  "
    strSQL = strSQL & "     FROM   " & P_LIB & ".SRAP01 "
    strSQL = strSQL & "     WHERE  RADLT=''  "
    strSQL = strSQL & "  ) AS RA ON RA.RATNO = X.JPTNO AND RA.RAENO=X.JPSWK AND RA.RACD1=X.JPHC4 "
    strSQL = strSQL & " LEFT JOIN  "
    strSQL = strSQL & "  ( SELECT  RATNO,RAENO,RACD1,RAPTN AS PTN_T "
    strSQL = strSQL & "     FROM   " & P_LIB & ".SRAP01 "
    strSQL = strSQL & "     WHERE  RADLT=''  "
    strSQL = strSQL & "  ) AS RA2 ON RA2.RATNO = X.JPTNO AND RA2.RAENO=' ' AND RA2.RACD1=' ' "
    strSQL = strSQL & " ORDER BY JPTNO "
    Debug.Print strSQL
    RS.Open strSQL, CN, adOpenStatic, adLockReadOnly
    Application.StatusBar = "データをシートに設定しています．．．"
    Do While Not RS.EOF
        If NVL(RTrim(RS("TEME1"))) = "福岡通過センター" Then
        Else
            行 = 行 + 1
            Cells(行, 2) = RS("TNO")
            Cells(行, 3) = RTrim(RS("TNM")) & " (" & RTrim(RS("TEME1")) & ")"
            Cells(行, 4) = "'" & RTrim(RS("SNO"))                               '2018/05/09　ゼロサプレス対応
            Cells(行, 5) = RS("JUCHU_SU")
            Cells(行, 6) = RS("SHUKKA_SU")
            Cells(行, 10) = "1"                                                 '2017/04/03 Add 出荷期限パターン
            If Trim(NVL(RS("PTN_T"))) <> "" Then Cells(行, 10) = RS("PTN_T")    ' 特約店共通のパターン
            If Trim(NVL(RS("PTN_N"))) <> "" Then Cells(行, 10) = RS("PTN_N")    ' 納品先別　　　〃
        End If
        RS.MoveNext
    Loop
    出荷先_最終行 = 行

    'ＲＳクローズ、ＤＢ切断
    RS.Close:    Set RS = Nothing
    CN.Close:    Set CN = Nothing

    '見ためを整える
    For 行 = 出荷先_行頭 To 出荷先_最終行
        '出荷済
        If Cells(行, 5) <= Cells(行, 6) Then
            Range(Cells(行, 5), Cells(行, 終了列)).Interior.Color = RGB(192, 192, 192) '灰色
        End If
    Next

    Cells(2, 7).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Application.StatusBar = False

End Sub

