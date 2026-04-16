Option Explicit

'                       2016/11/30　項目追加(作成ユーザ、仕分区分、汎用CD4、運送会社変更後、注文数)
'SQL文生成(追加)
Public Function InsertSQL(ByRef 出荷Rec As 出荷Record) As String
    Dim strSQL      As String

    strSQL = ""
    strSQL = strSQL & "INSERT INTO " & P_LIB & ".SZSP01"
    strSQL = strSQL & " ("
    strSQL = strSQL & " ZSDLT"
    strSQL = strSQL & ",ZSCNTH"
    strSQL = strSQL & ",ZSCTIM"
    strSQL = strSQL & ",ZSCUSR"
    strSQL = strSQL & ",ZSCPGM"
    strSQL = strSQL & ",ZSSDT"
    strSQL = strSQL & ",ZSNDT"
    strSQL = strSQL & ",ZSTNO"
    strSQL = strSQL & ",ZSSNO"
    strSQL = strSQL & ",ZSSGY"
    strSQL = strSQL & ",ZSDPK"
    strSQL = strSQL & ",ZSHNO"
    strSQL = strSQL & ",ZSJAN"
    strSQL = strSQL & ",ZSTNI"
    strSQL = strSQL & ",ZSLMT"
    strSQL = strSQL & ",ZSSRY"
    strSQL = strSQL & ",ZSYUCD"
    strSQL = strSQL & ",ZSYUCA"
    strSQL = strSQL & ",ZSSWK"
    strSQL = strSQL & ",ZSHC4"
    strSQL = strSQL & ",ZSKSU"
    strSQL = strSQL & ",ZSLOT"
    strSQL = strSQL & ",ZSSSTF"   '車両積荷前衛生点検（追加）
    strSQL = strSQL & ",ZSIDJK"   '逸脱事項（追加）
    strSQL = strSQL & ") VALUES ("
    strSQL = strSQL & " ''"                                             '削除フラグ
    strSQL = strSQL & ",TO_CHAR(current timestamp, 'YYYYMMDD')"         '作成日時
    strSQL = strSQL & ",TO_CHAR(current timestamp, 'HH24MISS')"         '作成時刻
    strSQL = strSQL & ",'" & P_USER & "'"                               '作成ユーザ(10桁)　SUPP01.USR=8桁なので溢れないはず
    strSQL = strSQL & ",'" & P_PGM & "'"                                '作成プログラム
    With 出荷Rec
        strSQL = strSQL & ", " & Format(.出荷日付, "yyyymmdd")          '出荷日付
        strSQL = strSQL & ", " & Format(.納品日付, "yyyymmdd")          '納品日付
        strSQL = strSQL & ",'" & .出荷先CD & "'"                        '出荷先CD
        strSQL = strSQL & ",'" & .伝票NO & "'"                          '伝票NO
        strSQL = strSQL & ", " & Val(.行NO)                             '行NO
        strSQL = strSQL & ",'" & .伝票区分 & "'"                        '伝票区分
        strSQL = strSQL & ",'" & .販売品番 & "'"                        '販売品番
        strSQL = strSQL & ",'" & .JAN & "'"                             'JAN
        strSQL = strSQL & ", " & Val(.単位)                             '単位
        strSQL = strSQL & ", " & Format(.賞味期限, "yyyymmdd")          '賞味期限
        strSQL = strSQL & ", " & Val(.出荷数量)                         '出荷数量
        strSQL = strSQL & ",'" & .運送会社CD & "'"                      '運送会社
        strSQL = strSQL & ",'" & .運送会社CD2 & "'"                     '運送会社変更後     '2016/11/30 Add
        strSQL = strSQL & ",'" & .仕分区分 & "'"                        '仕分区分           '2016/11/30 Add
        strSQL = strSQL & ",'" & .汎用CD4 & "'"                         '汎用CD4            '2016/11/30 Add
        strSQL = strSQL & ", " & Val(.注文数量)                         '注文数             '2016/11/30 Add
        strSQL = strSQL & ",'" & .ロットNO & "'"                        'ロット
        strSQL = strSQL & ", " & .車両積荷前衛生点検                     '1/0（11列目の値）
        strSQL = strSQL & ", '" & .逸脱事項 & "'"      '12列目のテキスト
    End With
    strSQL = strSQL & ")"
    Debug.Print strSQL
    InsertSQL = strSQL

End Function

'                       2016/11/30　「運送会社変更後」専用
'SQL文生成(更新)
Public Function UpdateSQL(ByRef 出荷Rec As 出荷Record) As String
    Dim strSQL      As String
    With 出荷Rec
        strSQL = ""
        strSQL = strSQL & "UPDATE " & P_LIB & ".SZSP01 "
        strSQL = strSQL & " SET "
        strSQL = strSQL & "  ZSUNTH =TO_CHAR(current timestamp, 'YYYYMMDD') "
        strSQL = strSQL & " ,ZSUTIM =TO_CHAR(current timestamp, 'HH24MISS') "
        strSQL = strSQL & " ,ZSUUSR ='" & P_USER & "'"
        strSQL = strSQL & " ,ZSUPGM ='" & P_PGM & "'"
        strSQL = strSQL & " ,ZSYUCA ='" & .運送会社CD2 & "'"
        strSQL = strSQL & " ,ZSSSTF = " & .車両積荷前衛生点検             '1/0（11列目の値）
        strSQL = strSQL & " ,ZSIDJK = '" & .逸脱事項 & "'"
        strSQL = strSQL & " WHERE ZSDLT='' "
        strSQL = strSQL & "   AND ZSSNO='" & .伝票NO & "'"
    End With
    Debug.Print strSQL
    UpdateSQL = strSQL

End Function

' 車両積荷前衛生点検・逸脱事項のみ更新
Public Function UpdateHygieneSQL(ByRef 出荷Rec As 出荷Record) As String
    Dim strSQL As String
    With 出荷Rec
        strSQL = ""
        strSQL = strSQL & "UPDATE " & P_LIB & ".SZSP01 "
        strSQL = strSQL & " SET "
        strSQL = strSQL & "  ZSSSTF = " & .車両積荷前衛生点検
        strSQL = strSQL & " ,ZSIDJK = '" & .逸脱事項 & "'"
        strSQL = strSQL & " WHERE ZSDLT='' "
        strSQL = strSQL & "   AND ZSSNO='" & .伝票NO & "'"
        strSQL = strSQL & "   AND ZSSGY=" & Val(.行NO)
    End With
    Debug.Print strSQL
    UpdateHygieneSQL = strSQL
End Function