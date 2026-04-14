Public Const P_PGM = "ShukkaXLS "
Public Const PB_SERVER = "M:\【包装出荷実績】\"

'接続文字列
Public Const P_接続文字列 = "DSN=ISA011;UID=SYSTEM;PWD=FJPN2480;"   'HONSHA
Public Const P_LIB = "LIBSMF17T"

'User.iniから
Public P_USER           As String
Public P_権限           As String

'ブック終了制御用
Public P_終了ボタン押下 As Boolean

'出荷先リストシートから
Public P_専用伝票NO     As String   '専用伝票NO(SNO)
Public P_運送会社CD     As String   '運送会社(JPHC9＋JPHC0)
Public P_運送会社NM     As String   '
Public P_出荷YMD        As Date     '出荷日
Public P_出荷期限KB     As String   '出荷期限パターン("1"or"2")

'明細シートから
Public P_行NO           As String   'frmZaikoへ渡す

'在庫一覧画面(frmZaiko)から
Public P_引当在庫メモ   As String   '明細シートへ渡す
Public P_引当更新       As Boolean  '明細シートへ渡す

'カレンダ画面から
Public P_カレンダ日付   As Date     '

'テンキー画面から
Public P_InputTenKey    As String

'                                   '2016/11/30　項目追加(仕分区分、汎用CD4、運送会社変更後、注文数)
'出荷TBL更新用ワーク(SZSP01)
Public P_出荷Rec As 出荷Record
Public Type 出荷Record
    出荷日付    As Date
    納品日付    As Date
    出荷先CD    As String
    伝票NO      As String
    行NO        As String
    伝票区分    As String
    販売品番    As String
    生産品番    As String
    JAN         As String
    単位        As String
    賞味期限    As Date
    出荷数量    As String
    運送会社CD  As String
    仕分区分    As String
    汎用CD4     As String
    注文数量    As String
    運送会社CD2 As String
    ロットNO    As String
    車両積荷前衛生点検 As Integer   'ZSSSTF 追加
    逸脱事項    As String           'ZSIDJK 追加

End Type

'変数を保存する
Public Function Set共通変数() As Long

    '出荷先リスト
    st01List.Cells(1, 101) = "出荷先_最終行":    st01List.Cells(1, 102) = 出荷先_最終行
    st01List.Cells(2, 101) = "専用伝票NO":       st01List.Cells(2, 102) = "'" & P_専用伝票NO '2018/05/09　ゼロサプレス対応
    st01List.Cells(3, 101) = "運送会社":         st01List.Cells(3, 102) = "'" & P_運送会社CD
                                                 st01List.Cells(3, 103) = P_運送会社NM
    st01List.Cells(4, 101) = "出荷YMD":          st01List.Cells(4, 102) = P_出荷YMD
    
    '明細
    st02Meisai.Cells(1, 101) = "明細_最終行":    st02Meisai.Cells(1, 102) = 明細_最終行
    st02Meisai.Cells(2, 101) = "行NO":           st02Meisai.Cells(2, 102) = P_行NO

    '引当ワーク
    st02Hikiate.Cells(1, 101) = "引当_最終行":   st02Hikiate.Cells(1, 102) = 引当_最終行

End Function

'変数を取得する
Public Function Get共通変数() As Long

    '出荷先リスト
    出荷先_最終行 = Val(st01List.Cells(1, 102))
    P_専用伝票NO = st01List.Cells(2, 102)
    P_運送会社CD = st01List.Cells(3, 102):        P_運送会社NM = st01List.Cells(3, 103)
    P_出荷YMD = CDate(st01List.Cells(4, 102))
    
    '明細
    明細_最終行 = Val(st02Meisai.Cells(1, 102))
    P_行NO = Val(st02Meisai.Cells(2, 102))

    '引当ワーク
    引当_最終行 = Val(st02Hikiate.Cells(1, 102))

    'Iniファイル
    If P_USER = "" Then Call ReadUserIni(Application.ThisWorkbook.Path & "\User.ini")
End Function
'※Val関数の戻り値はdouble


Public Function Get賞味期限fromロット(ByVal ロット As String) As Date
    Get賞味期限fromロット = 0

    If Len(RTrim(ロット)) <> 10 Then Exit Function
    Get賞味期限fromロット = CDate(Format(Mid(ロット, 1, 8), "0000/00/00"))

End Function

Public Function Getバッチ数fromロット(ByVal ロット As String) As String
    Getバッチ数fromロット = ""

    If Len(RTrim(ロット)) <> 10 Then Exit Function
    Getバッチ数fromロット = Mid(ロット, 9, 2)

End Function
