Option Explicit

'                       2016/11/30　項目追加(作成ユーザ、更新ユーザ)
'SQL文生成(更新)
Public Function UpdateSQL_在庫(ByRef 出荷Rec As 出荷Record) As String
    Dim strSQL      As String
    
    With 出荷Rec
        strSQL = ""
        strSQL = strSQL & "UPDATE " & P_LIB & ".SSZP01 "
        strSQL = strSQL & " SET "
        strSQL = strSQL & "  SZUNTH =TO_CHAR(current timestamp, 'YYYYMMDD') "
        strSQL = strSQL & " ,SZUTIM =TO_CHAR(current timestamp, 'HH24MISS') "
        strSQL = strSQL & " ,SZUUSR ='" & P_USER & "'"
        strSQL = strSQL & " ,SZUPGM ='" & P_PGM & "'"
        strSQL = strSQL & " ,SZSRY  = SZSRY - " & Val(.出荷数量)                '在庫数量
        strSQL = strSQL & " ,SZHNO  ='" & .販売品番 & "'"
        strSQL = strSQL & " WHERE SZDLT='' "
        strSQL = strSQL & "   AND SZLOT='" & .ロットNO & "'"
        strSQL = strSQL & "   AND SZSNO='" & .生産品番 & "'"
    End With
    Debug.Print strSQL
    UpdateSQL_在庫 = strSQL

End Function

'SQL文生成(追加)
Public Function InsertSQL_在庫(ByRef 出荷Rec As 出荷Record) As String
    Dim strSQL      As String

    With 出荷Rec
        strSQL = ""
        strSQL = strSQL & "INSERT INTO " & P_LIB & ".SSZP01"
        strSQL = strSQL & " ("
        strSQL = strSQL & " SZDLT"
        strSQL = strSQL & ",SZCNTH"
        strSQL = strSQL & ",SZCTIM"
        strSQL = strSQL & ",SZCUSR"
        strSQL = strSQL & ",SZCPGM"
        strSQL = strSQL & ",SZLNO"
        strSQL = strSQL & ",SZSNO"
        strSQL = strSQL & ",SZHNO"
        strSQL = strSQL & ",SZJAN"
        strSQL = strSQL & ",SZLMT"
        strSQL = strSQL & ",SZSRY"
        strSQL = strSQL & ",SZCRY"
        strSQL = strSQL & ",SZCDT"
        strSQL = strSQL & ",SZLOT"
        strSQL = strSQL & ",SZFG1"
        strSQL = strSQL & ",SZFG2"
        strSQL = strSQL & ") VALUES ("
        strSQL = strSQL & " ''"                                             '削除フラグ
        strSQL = strSQL & ",TO_CHAR(current timestamp, 'YYYYMMDD')"         '作成日時
        strSQL = strSQL & ",TO_CHAR(current timestamp, 'HH24MISS')"         '作成時刻
        strSQL = strSQL & ",'" & P_USER & "'"                               '作成ユーザ
        strSQL = strSQL & ",'" & P_PGM & "'"                                '作成プログラム
        strSQL = strSQL & ",'   '"                                          'ラインNo.   ※空白固定
        strSQL = strSQL & ",'" & .生産品番 & "'"                            '生産品番
        strSQL = strSQL & ",'" & .販売品番 & "'"                            '販売品番
        strSQL = strSQL & ",'" & .JAN & "'"                                 'JAN         ※空白固定
        strSQL = strSQL & ", " & Format(.賞味期限, "yyyymmdd")              '賞味期限
        strSQL = strSQL & ", " & Val(.出荷数量 * -1)                        '現在庫数量
        strSQL = strSQL & ", 0"                                             '在庫調整数
        strSQL = strSQL & ", 0"                                             '在庫調整日
        strSQL = strSQL & ",'" & .ロットNO & "'"                            'LOT
        strSQL = strSQL & ",' '"
        strSQL = strSQL & ",' '"
        strSQL = strSQL & ")"
    End With
    Debug.Print strSQL
    InsertSQL_在庫 = strSQL

End Function

