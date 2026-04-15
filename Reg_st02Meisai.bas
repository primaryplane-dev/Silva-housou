Option Explicit

Private Const 列_引当 = 9
Private Const 列_チェック = 10

'②在庫の引当てを選択する
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim 行          As Long
    Dim 列          As Integer

    '画面を閉じると値が消えるので再取得
    Call Get共通変数

    If Not (Target.Column = 列_引当 Or Target.Column = 列_チェック) Then Exit Sub
    If Target.Row < 明細_行頭 Then Exit Sub
    If Target.Row > 明細_最終行 Then Exit Sub
    If Target.Count > 1 Then Exit Sub                  '複数セル
    If Cells(Target.Row, 8).Value = 0 Then Exit Sub    '注文数ゼロ
    行 = Target.Row
    列 = Target.Column

    Select Case 列
    Case 列_引当
        P_行NO = Cells(行, 2).Value
        If Val(P_行NO) = 0 Then Exit Sub
        Call Set共通変数
        '在庫一覧画面を表示する
        Load frmZaiko
        frmZaiko.Show
        '更新されたら反映する
        If P_引当更新 = True Then
            st02Meisai.Select
            Cells(行, 9).Value = P_引当在庫メモ
            Cells(行, 10).Value = "引当する"
            Cells(行, 列_チェック).Interior.Color = RGB(255, 153, 204)  '桃色
            Rows(行).AutoFit
            ' 11列目: 車両積荷前衛生点検（〇→1、×→0）
            Dim tmpVal As String
            tmpVal = Trim(Cells(行, 11).Value)
            Dim tmpHikiate As Variant
            If tmpVal = "〇" Then
                Cells(行, 11).Value = 1
                tmpHikiate = 1
            ElseIf tmpVal = "×" Then
                Cells(行, 11).Value = 0
                tmpHikiate = 0
            Else
                tmpHikiate = ""
            End If
            ' 18列目: 車両積荷前衛生点検（1/0）をst02Hikiateへ転記
            st02Hikiate.Cells(行, 18).Value = tmpHikiate
            ' 19列目: 逸脱事項（AS項目:ZSIDJK／st02Hikiateへそのまま転記）
            st02Hikiate.Cells(行, 19).Value = Cells(行, 12).Value
            Call Set共通変数
        End If
    
    Case 列_チェック
        'チェック欄 切換え
        If Cells(行, 列).Value = "未処理" Then
            Cells(行, 列).Value = "引当する"
            Range(Cells(行, 列), Cells(行, 列)).Interior.Color = RGB(255, 153, 204)  '桃色
        ElseIf Cells(行, 列).Value = "引当する" Then
            Cells(行, 列).Value = "未処理"
            Range(Cells(行, 列), Cells(行, 列)).Interior.ColorIndex = xlNone
        End If
    End Select

End Sub

'出荷日を変更する
Private Sub cmd出荷日_Click()
    P_カレンダ日付 = Cells(3, 3)
    frmカレンダ.Show
    If P_カレンダ日付 > 0 Then Cells(3, 3) = P_カレンダ日付
End Sub

Private Sub cmdリストへ戻る_Click()

    If 変更チェック = "変更あり" Then
        If MsgBox("変更内容を破棄してよろしいですか", vbOKCancel) = vbCancel Then Exit Sub
    End If

    st01List.Select
End Sub

Private Sub cmd確定する_Click()
    Dim CN                  As Object
    Dim 行                  As Long
    Dim data行              As Long
    Dim 列                  As Integer
    Dim i                   As Integer
    Dim 出荷Rec             As 出荷Record
    Dim ret                 As Boolean
    Dim 更新FLG             As Boolean
    Dim ErrFLG              As Boolean
    
    'チェック
    Call Get共通変数
    If 明細_最終行 = 0 Then Exit Sub
    If 引当_最終行 = 0 Then Exit Sub
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.StatusBar = "データを登録しています．．．"
    
    Set CN = CreateObject("ADODB.Connection")
    CN.CursorLocation = adUseClient
    CN.Open P_接続文字列
    If Not P_LIB = "LIBSMF17T" Then
        CN.BeginTrans
    End If

    'ヘッダ情報を格納する
    With 出荷Rec
        .出荷日付 = Cells(3, 3).Value
        .納品日付 = Cells(4, 3).Value
        .出荷先CD = Cells(5, 3).Value
        .伝票NO = Cells(4, 8).Value
        .JAN = 0
        .運送会社CD = P_運送会社CD
        If st02Meisai.cbo運送会社変更後.Text <> "" Then .運送会社CD2 = st02Meisai.cbo運送会社変更後.Value   '2016/11/30 Add
        If Len(Cells(5, 4).Value) >= 4 Then .仕分区分 = Mid(Cells(5, 4).Value, 3, 2)                        '2016/11/30 Add
        If Len(Cells(5, 4).Value) >= 1 Then .汎用CD4 = Mid(Cells(5, 4).Value, 1, 1)                         '2016/11/30 Add
    End With

    '「引当する」を指定した行を更新対象にする
    ErrFLG = False
    更新FLG = False
    For 行 = 明細_行頭 To 明細_最終行
        If Cells(行, 列_チェック).Value = "引当する" Then
            With 出荷Rec
                .行NO = Cells(行, 2).Value
                .伝票区分 = Cells(行, 3).Value
                .販売品番 = Cells(行, 4).Value
                '商品ひとつにつき、複数の在庫を引当てる(ロットが異なる)
                For data行 = 引当_行頭 To 引当_最終行
                    If .行NO = st02Hikiate.Cells(data行, 3).Value Then
                        Select Case st02Hikiate.Cells(data行, 15).Value
                        Case "*", "**", "切*"
                            .単位 = st02Hikiate.Cells(data行, 8).Value
                            .注文数量 = st02Hikiate.Cells(data行, 10).Value     '2016/11/30 Add
                            .出荷数量 = st02Hikiate.Cells(data行, 14).Value
                            .ロットNO = st02Hikiate.Cells(data行, 16).Value
                            .賞味期限 = Get賞味期限fromロット(.ロットNO)
                            .生産品番 = st02Hikiate.Cells(data行, 12).Value
                            ' 11列目: 車両積荷前衛生点検（1/0）を直接セット
                            .車両積荷前衛生点検 = st02Hikiate.Cells(data行, 18).Value
                            ' 12列目: 逸脱事項（フリー入力）を直接セット
                            .逸脱事項 = st02Hikiate.Cells(data行, 19).Value
                            If DB更新(出荷Rec, CN) = True Then
                                更新FLG = True
                            Else
                                ErrFLG = True
                                Exit For
                            End If
                        End Select
                    End If
                Next
                If ErrFLG = True Then Exit For
            End With
        End If
    Next
    
    If st02Meisai.cbo運送会社変更後.Text <> "" Then         '2016/11/30 Add 全行同じなので最終行の値を使う
        If DB更新_運送会社(出荷Rec, CN) = True Then
            更新FLG = True
        Else
            ErrFLG = True
        End If
    End If
    
    If Not P_LIB = "LIBSMF17T" Then
        If ErrFLG = True Then
            CN.RollbackTrans
        Else
            CN.CommitTrans
        End If
    End If
    CN.Close:    Set CN = Nothing
    
    'シートの再描画                                         '2017/05/03 Upd 56 明細シートにとどまる
    Call 出荷先リスト表示
    st01List.Select
    Call st01List.カレント行移動
    Call Create在庫引当ワーク
    Call 明細表示
    Call Set共通変数
    st02Meisai.Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False

End Sub

Private Function 変更チェック() As String
    Dim 行                  As Long

    '初期値
    変更チェック = "変更なし"

    For 行 = 明細_行頭 To 明細_最終行
        If st02Meisai.Cells(行, 10).Value = "引当する" Then
            変更チェック = "変更あり"
            Exit Function
        End If
    Next

End Function
