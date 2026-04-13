Option Explicit

Private Const レ点 = "a"            'フォントを「Marlett」にすること
Private Const 最終列 = 4

'①一覧を表示する
Private Sub cmdリスト表示_Click()
    Application.EnableEvents = False
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    Call 一覧表示
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault
End Sub

'②行を選択する　　　→商品、賞味期限が決まる
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim 行          As Long

    リスト_最終行 = Cells(5, 1).Value
    If Target.Column > 4 Then Exit Sub
    If Target.Row < リスト_開始行 Then Exit Sub
    If Target.Row > リスト_最終行 Then Exit Sub
    If Target.Count > 1 Then Exit Sub                   '複数セル
    If Cells(Target.Row, 2).Value = "" Then Exit Sub    '空行
    行 = Target.Row

    '行の選択/選択解除
    If Cells(行, 1).Value = レ点 Then
        Cells(行, 1).Value = ""
        Range(Cells(行, 1), Cells(行, 最終列)).Interior.ColorIndex = xlNone
    Else
        '他の行をクリアする
        Range(Cells(6, 1), Cells(リスト_最終行, 1)).Value = ""
        Range(Cells(6, 1), Cells(リスト_最終行, 最終列)).Interior.ColorIndex = xlNone
        '選択行をピンクにする
        Cells(行, 1).Value = レ点
        Range(Cells(行, 1), Cells(行, 最終列)).Interior.Color = RGB(255, 153, 204)
    End If

End Sub

'③明細シートへ
Private Sub cmd明細表示_Click()
    Dim 行          As Long
    Dim 選択行      As Long

    '入力チェック
    リスト_最終行 = Cells(5, 1).Value
    If リスト_最終行 = 0 Then Exit Sub

    '選択中の行を探す
    For 行 = リスト_開始行 To リスト_最終行
        If Cells(行, 1).Value = レ点 Then
            選択行 = 行
            Exit For
        End If
    Next
    If 選択行 = 0 Then Exit Sub
    If Val(Cells(選択行, 2).Value) = 0 Then Exit Sub
    If Val(Cells(選択行, 4).Value) = 0 Then Exit Sub
    
    '条件を保存する
    P_品番 = Cells(選択行, 2).Value
    P_賞味期限 = Cells(選択行, 4).Value
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.StatusBar = "データをシートに設定しています．．．"
    
    '明細シートへ
    Call 明細表示

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    st02Meisai.Activate
End Sub

'条件入力
Private Sub cmd日付選択_Click()
    'カレンダ画面で日付を選択する
    P_カレンダ日付 = P_賞味期限日
    frmカレンダ.Show
    P_賞味期限日 = P_カレンダ日付
    If P_賞味期限日 = 0 Then
        Cells(2, 3).Value = ""
    Else
        Cells(2, 3).Value = P_賞味期限日
    End If
End Sub
Private Sub cmd賞味期限クリア_Click()
    P_賞味期限日 = 0
    Cells(2, 3).Value = ""
    
'    Call 一覧クリア '2017/05/01 Delete
'    Cells(2, 3).Select '2017/05/01 Delete
End Sub

Private Sub cmd商品検索_Click()
    Dim i                   As Integer
    Dim strWK               As String
    
    '品名カナ検索画面で品番を選択する(複数)
    P_検索FLG = False
    Load FrmSearchName
    FrmSearchName.Show

    If P_検索FLG = True Then
        If UBound(P_製品) > 0 Then
            For i = 1 To UBound(P_製品)
                If i > 1 Then strWK = strWK & ", "
                strWK = strWK & P_製品(i).CD
            Next
        End If
        Cells(3, 3).Value = strWK '2017/05/01 Update
    End If
    'Cells(3, 3).Value = strWK
End Sub
Private Sub cmd商品クリア_Click()
    ReDim P_製品(0)
    Cells(3, 3).Value = ""
    
'    Call 一覧クリア '2017/05/01 Delete
'    Cells(3, 3).Select '2017/05/01 Delete
End Sub

'
'終了
Private Sub cmd終了_Click()
    P_終了ボタン押下 = True
    Application.Quit
    ActiveWorkbook.Close False
End Sub


