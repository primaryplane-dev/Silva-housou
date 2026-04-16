Option Explicit

Private Const レ点 = "a"            'フォントを「Marlett」にすること
Private Const 最終列 = 4

Private Sub cmd出荷日_Click()
  
    '出荷日を指定する(カレンダ画面)
    P_カレンダ日付 = Cells(3, 3)
    frmカレンダ.Show
    If P_カレンダ日付 > 0 Then Cells(3, 3) = P_カレンダ日付
    
End Sub

'①出荷データを抽出する
Private Sub cmdリスト表示_Click()

    '入力チェック
    If st01List.cbo運送会社.ListIndex <= 0 Then Call MsgBox("運送会社を選択してください"):  Exit Sub
    If IsDate(Cells(3, 3)) = False Then Call MsgBox("出荷日を選択してください"):   Exit Sub
     
    '条件を保存
    P_運送会社CD = st01List.cbo運送会社.Value
    P_運送会社NM = st01List.cbo運送会社.Text
    P_出荷YMD = CDate(Cells(3, 3))
    
    'リストを再表示する
    Application.EnableEvents = False
    
    Call 出荷先リスト表示
    Call Set共通変数
    
    Application.EnableEvents = True

End Sub

'②行を選択する　　　→専用伝票NOが決まる(出荷先も決まる)
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim 行          As Long

   'If Target.Column <> 1 Then Exit Sub                 '2016/11/30 Upd
    If Target.Column > 最終列 Then Exit Sub             '
    If Target.Row < 出荷先_行頭 Then Exit Sub
    If Target.Count > 1 Then Exit Sub                   '複数セル
    If Cells(Target.Row, 2).Value = "" Then Exit Sub    '空行
    行 = Target.Row

    '行の選択/選択解除
    If Cells(行, 1).Value = レ点 Then
        Cells(行, 1).Value = ""
        Range(Cells(行, 1), Cells(行, 最終列)).Interior.ColorIndex = xlNone
    Else
        '他の行をクリアする
        Range(Cells(6, 1), Cells(100, 1)).Value = ""
        Range(Cells(6, 1), Cells(100, 最終列)).Interior.ColorIndex = xlNone
        '選択行をピンクにする
        Cells(行, 1).Value = レ点
        Range(Cells(行, 1), Cells(行, 最終列)).Interior.Color = RGB(255, 153, 204)
    End If

End Sub

'③明細シートに移動する
Private Sub cmd出荷入力_Click()
    Dim 行          As Long
    Dim 選択行      As Long

    '入力チェック
    Call Get共通変数
    If P_運送会社CD = "" Then Exit Sub

    '選択中の行を探す
    選択行 = 0
    For 行 = 出荷先_行頭 To 出荷先_最終行
        If Cells(行, 1).Value = レ点 Then
            選択行 = 行
            Exit For
        End If
    Next
    If 選択行 = 0 Then
        MsgBox "出荷先リストから行を選択してください"
        Exit Sub
    End If
    If Val(Cells(選択行, 2).Value) = 0 Then
        MsgBox "選択行にデータがありません"
        Exit Sub
    End If
    
    '条件を保存する
    P_専用伝票NO = Cells(選択行, 4).Value
    P_出荷期限KB = Cells(選択行, 10).Value

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.StatusBar = "データをシートに設定しています．．．"
    
    '明細シートへ
    Call Create在庫引当ワーク
    Call 明細表示
    Call Set共通変数

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    st02Meisai.Activate
End Sub

Public Sub カレント行移動()
    Dim 行          As Long

    '専用伝票NOを元にカレント行を移動する
    For 行 = 出荷先_行頭 To 出荷先_最終行
        If Cells(行, 4).Value = P_専用伝票NO Then
            Cells(行, 1).Value = レ点
            Range(Cells(行, 1), Cells(行, 最終列)).Interior.Color = RGB(255, 153, 204)
            Exit For
        End If
    Next
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 行

End Sub

'別ブック起動
Private Sub cmd在庫調整へ_Click()
    Call M_sbClickButton("包装在庫調整.xls")
End Sub

'終了
Private Sub cmd終了_Click()
    P_終了ボタン押下 = True
    Application.Quit
    ActiveWorkbook.Close False
End Sub


