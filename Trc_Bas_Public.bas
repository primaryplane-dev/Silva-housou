Option Explicit

Public Const P_接続文字列 = "DSN=ISA011;UID=SYSTEM;PWD=FJPN2480;"   'HONSHA
Public Const P_LIB = "LIBSMF17"


'リストシートから(明細シート表示の条件)
Public P_賞味期限       As String
Public P_品番           As String


'カレンダ画面から
Public P_カレンダ日付   As Date
Public P_賞味期限日     As Date         'リストシート表示の条件

'品名検索画面から
Public P_製品()         As udt名称      'リストシート表示の条件
Public Type udt名称
    CD                  As String
    NM                  As String
End Type
Public P_検索FLG        As Boolean

'ブック終了制御用
Public P_終了ボタン押下 As Boolean



