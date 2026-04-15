Option Explicit

Private Sub Workbook_Open()

    ThisWorkbook.Activate

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    '変数クリア
    P_終了ボタン押下 = False

    'INIを読む
    Call ReadUserIni(Application.ThisWorkbook.Path & "\User.ini")

    ' ブックを保存しない前提だが念のため。
    'シートをクリアする
    ' 11列目（車両積荷前衛生点検）はst02Hikiateシートで〇/×入力→明細シートで1/0に変換してAS連携
    ' 12列目（逸脱事項）はst02Hikiateシートでフリー入力→明細シート・AS連携もこの値を使う
  
    Call 出荷先リストクリア
    Call 明細クリア
    Call 在庫引当クリア
    If P_権限 = "1" Then st02Meisai.cmd確定する.Enabled = True

    'コンボボックスの選択リスとを作る
    Call Create運送会社リスト
    st01List.cbo運送会社.ListIndex = 0
    st01List.Select

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If P_終了ボタン押下 = False Then
        MsgBox "閉じるボタンは使用できません。" & vbCrLf & "リストシートの終了ボタンで終了してください。", vbCritical
        Cancel = True
    End If
End Sub
