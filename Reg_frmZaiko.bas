Option Explicit

Private data行_先頭 As Long
Private 引当Rec()           As 引当Record

Private Sub UserForm_Initialize()
    Dim data行              As Long
    Dim 行                  As Integer

    P_引当更新 = False

    '在庫引当シートのデータを抽出(キー：行NO)
    ReDim 引当Rec(0)
    For data行 = 引当_行頭 To 引当_最終行
        If st02Hikiate.Cells(data行, 3) = P_行NO Then
            行 = 行 + 1:    ReDim Preserve 引当Rec(行)
            With 引当Rec(行)
                '注文
                .伝票NO = st02Hikiate.Cells(data行, 2)
                .行NO = st02Hikiate.Cells(data行, 3)
                .伝票区分 = st02Hikiate.Cells(data行, 4)
                .販売品番 = st02Hikiate.Cells(data行, 5)
                .販売品名 = st02Hikiate.Cells(data行, 6)
                .入数 = st02Hikiate.Cells(data行, 7)
                .単位 = st02Hikiate.Cells(data行, 8)
                .単位名 = st02Hikiate.Cells(data行, 9)
                .注文数 = st02Hikiate.Cells(data行, 10)
                '出荷/在庫
                .販売品番2 = st02Hikiate.Cells(data行, 11)
                .生産品番 = st02Hikiate.Cells(data行, 12)
                .在庫数 = st02Hikiate.Cells(data行, 13)
                .出荷数 = st02Hikiate.Cells(data行, 14)
                .区分 = st02Hikiate.Cells(data行, 15)
                .ロット = st02Hikiate.Cells(data行, 16)
                .賞味期限 = Get賞味期限fromロット(.ロット)
                .バッチNO = Getバッチ数fromロット(.ロット)
                .出庫期限 = st02Hikiate.Cells(data行, 17)
            End With
        End If
    Next
    
    Call 在庫データ表示

End Sub

Private Sub 在庫データ表示()
    Dim 行                  As Integer
    Dim str行               As String
    Dim data行              As Long
    Dim WK出荷済数          As Long
    Dim WK割当数計          As Long

    For data行 = 1 To UBound(引当Rec)
        With 引当Rec(data行)
            'ヘッダを表示する
            If 行 = 0 Then
                data行_先頭 = data行
                lbl販売品番.Caption = .販売品番
                lbl販売品名.Caption = .販売品名
                lbl出荷数.Caption = .注文数
            End If
                
            '在庫情報を表示する(*:引当、**：手入力在庫、+：在庫、x：期限切れ在庫、切*：引当(期限切れ))
            Select Case .区分
            Case "*", "**", "+", "x", "切*"
                行 = 行 + 1:    str行 = Format(行, "00")
                If 行 <= 5 Then
                    Me.Controls("lbl賞味期限_" & str行).Caption = Format(.賞味期限, "yyyy/mm/dd")
                    Me.Controls("lblバッチ数_" & str行).Caption = .バッチNO
                    Me.Controls("lbl在庫数_" & str行).Caption = .在庫数
                    Me.Controls("lbl割当数_" & str行).Caption = .出荷数
                    Me.Controls("lbl区分_" & str行).Caption = .区分
                    Me.Controls("lbl引当説明_" & str行).Caption = 引当マーク文言変換(.区分)
                    If .出庫期限 = 0 Then
                        Me.Controls("lbl出庫期限_" & str行).Caption = ""            'データがない場合は空白　2016/10/27 榊原
                        Else
                        Me.Controls("lbl出庫期限_" & str行).Caption = Format(.出庫期限, "yyyy/mm/dd")
                    End If
                End If
            End Select
        
            '集計する
            Select Case .区分
            Case "確":              WK出荷済数 = WK出荷済数 + .出荷数
            Case "*", "**", "切*":  WK割当数計 = WK割当数計 + .出荷数
            End Select
        End With
    Next

    '合計欄を表示する
    lbl出荷済数.Caption = WK出荷済数
    lbl割当数計.Caption = WK割当数計

End Sub

Private Sub cmd引当を決定する_Click()
    Dim data行              As Long
    Dim 行                  As Integer
    Dim str行               As String
    Dim 追加行数            As Integer
    Dim 削除行数            As Integer
    Dim WK割当数            As Long
    Dim WK割当数計          As Long
    Dim WKメモ              As String
    
    '■在庫引当シートに書きなおす
    '（今回の引当内容を末尾に追加する）
    data行 = 引当_最終行
    For 行 = 1 To 5
        str行 = Format(行, "00")
        Select Case Me.Controls("lbl区分_" & str行).Caption
        Case "+", "*", "**", "x", "切*"
            If Me.Controls("lbl割当数_" & str行).Caption <> "" Then
                data行 = data行 + 1
                追加行数 = 追加行数 + 1
                '注文
                st02Hikiate.Cells(data行, 2) = "'" & 引当Rec(1).伝票NO          '2018/05/09　ゼロサプレス対応
                st02Hikiate.Cells(data行, 3) = 引当Rec(1).行NO
                st02Hikiate.Cells(data行, 4) = 引当Rec(1).伝票区分
                st02Hikiate.Cells(data行, 5) = 引当Rec(1).販売品番
                st02Hikiate.Cells(data行, 6) = 引当Rec(1).販売品名
                st02Hikiate.Cells(data行, 7) = 引当Rec(1).入数
                st02Hikiate.Cells(data行, 8) = 引当Rec(1).単位
                st02Hikiate.Cells(data行, 9) = 引当Rec(1).単位名
                st02Hikiate.Cells(data行, 10) = 引当Rec(1).注文数
                '出荷/在庫
                st02Hikiate.Cells(data行, 11) = lbl販売品番.Caption
                st02Hikiate.Cells(data行, 12) = lbl販売品番.Caption     'lbl生産品番.Caption
                st02Hikiate.Cells(data行, 13) = Me.Controls("lbl在庫数_" & str行).Caption
                st02Hikiate.Cells(data行, 14) = Me.Controls("lbl割当数_" & str行).Caption
                st02Hikiate.Cells(data行, 15) = Me.Controls("lbl区分_" & str行).Caption
                st02Hikiate.Cells(data行, 16) = Format(CDate(Me.Controls("lbl賞味期限_" & str行).Caption), "yyyymmdd") _
                                              & Me.Controls("lblバッチ数_" & str行).Caption
                st02Hikiate.Cells(data行, 17) = Me.Controls("lbl出庫期限_" & str行).Caption
            End If
        End Select
    Next
    
    '変更なし。ここで終わり                                                 '2017/05/08 Upd No.61 割当数なしのとき実行時エラー
   'If 追加行数 = 0 Then Unload Me
    If 追加行数 = 0 Then Unload Me: Exit Sub
    
    '見ためを整える
    st02Hikiate.Activate
    Dim KEY As String
    Dim KEY_Z As String
    Range(Cells(引当_行頭, 2), Cells(data行, 17)).Borders.LineStyle = xlContinuous
    For data行 = 引当_行頭 To data行
        KEY = Cells(data行, 2) & Cells(data行, 3) & Cells(data行, 4) & Cells(data行, 5) & Cells(data行, 6) & Cells(data行, 7)
        If KEY = KEY_Z Then
            st02Hikiate.Range(st02Hikiate.Cells(data行, 2), st02Hikiate.Cells(data行, 10)).Font.Color = RGB(192, 192, 192)
        Else
            KEY_Z = KEY
        End If
    Next

    '既存のレコードを削除する
    For data行 = 引当_最終行 To 引当_行頭 Step -1
        If P_行NO = Cells(data行, 3) Then
            Select Case st02Hikiate.Cells(data行, 15)
            Case "+", "*", "**", "x", "切*"
                st02Hikiate.Rows(data行).Delete Shift:=xlUp
                削除行数 = 削除行数 + 1
            End Select
        End If
    Next
    引当_最終行 = 引当_最終行 + 追加行数 - 削除行数

    '■結果を戻す
    '引当メモを生成する
    For data行 = 1 To UBound(引当Rec)
        With 引当Rec(data行)
            If .区分 = "確" Then
                If WKメモ <> "" Then WKメモ = WKメモ & vbCrLf
                                                                            '2017/05/08 Upd No.62 バッチNoを削除
               'WKメモ = WKメモ & Right(String(6, " ") & .出荷数, 6) _
               '                & " (" & Format(.賞味期限, "yyyy/mm/dd") _
               '                       & "-" & .バッチNO & ") " _
               '                & .区分
                WKメモ = WKメモ & Right(String(6, " ") & .出荷数, 6) _
                                & " (" & Format(.賞味期限, "yyyy/mm/dd") & ") " _
                                & .区分
            End If
        End With
    Next
    For 行 = 1 To 5
        str行 = Format(行, "00")
        Select Case Me.Controls("lbl区分_" & str行).Caption
        Case "*", "**", "x", "切*"
            If Val(Me.Controls("lbl割当数_" & str行).Caption) <> 0 Then
                If WKメモ <> "" Then WKメモ = WKメモ & vbCrLf
                                                                            '2017/05/08 Upd No.62 バッチNoを削除
               'WKメモ = WKメモ & Right(String(6, " ") & Me.Controls("lbl割当数_" & str行).Caption, 6) _
               '                & " (" & Me.Controls("lbl賞味期限_" & str行).Caption _
               '                & "-" & Me.Controls("lblバッチ数_" & str行).Caption & ") " _
               '                & Me.Controls("lbl区分_" & str行).Caption
                WKメモ = WKメモ & Right(String(6, " ") & Me.Controls("lbl割当数_" & str行).Caption, 6) _
                                & " (" & Me.Controls("lbl賞味期限_" & str行).Caption & ") " _
                                & Me.Controls("lbl区分_" & str行).Caption
            End If
        End Select
    Next
    P_引当更新 = True
    P_引当在庫メモ = WKメモ
    Unload Me
    
End Sub

'キャンセルボタン
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'---------------------------------------
' 別の在庫
'---------------------------------------
Private Sub Input賞味期限(ByVal 行 As String)
    
    '区分が入っている行は抜ける
    Select Case Me.Controls("lbl区分_" & 行).Caption
    Case "", "**":  '入力可能
    Case Else:      Exit Sub
    End Select

    '賞味期限を入力する
    P_カレンダ日付 = 0
    If Me.Controls("lbl賞味期限_" & 行).Caption <> "" Then P_カレンダ日付 = Me.Controls("lbl賞味期限_" & 行).Caption
    frmカレンダ.Show
    If P_カレンダ日付 = 0 Then Exit Sub
    Me.Controls("lbl賞味期限_" & 行).Caption = Format(P_カレンダ日付, "yyyy/mm/dd")

    'バッチ数に初期値を設定する
    If Me.Controls("lblバッチ数_" & 行).Caption = "" Then
        Me.Controls("lblバッチ数_" & 行).Caption = "00"
    End If

    '区分を「引当済(別在庫)」にする
    Me.Controls("lbl区分_" & 行).Caption = "**"
    
End Sub

Private Sub lbl賞味期限_01_Click()
    Call Input賞味期限("01")
End Sub
Private Sub lbl賞味期限_02_Click()
    Call Input賞味期限("02")
End Sub
Private Sub lbl賞味期限_03_Click()
    Call Input賞味期限("03")
End Sub
Private Sub lbl賞味期限_04_Click()
    Call Input賞味期限("04")
End Sub
Private Sub lbl賞味期限_05_Click()
    Call Input賞味期限("05")
End Sub

Private Sub Input割当数(ByVal 行 As String)

    '出荷確定済の行は抜ける
    '賞味期限未入力の行は抜ける
    If Me.Controls("lbl区分_" & 行).Caption = "確" Then Exit Sub
    If Me.Controls("lbl賞味期限_" & 行).Caption = "" Then Exit Sub
    
    '割当数を入力する
    P_InputTenKey = 0
    If Val(Me.Controls("lbl割当数_" & 行).Caption) <> 0 Then
        P_InputTenKey = Me.Controls("lbl割当数_" & 行).Caption
    End If
    frmInputTenKey.Show
   'Me.Controls("lbl割当数_" & 行).Caption = Format(P_InputTenKey, "00")   '2017/05/08 Upd 前ゼロ詰めしない
    Me.Controls("lbl割当数_" & 行).Caption = Format(Val(P_InputTenKey), "0")

    '割当数計を再計算する
    lbl割当数計.Caption = Val(lbl割当数_01.Caption) + Val(lbl割当数_02.Caption) + Val(lbl割当数_03.Caption) _
                        + Val(lbl割当数_04.Caption) + Val(lbl割当数_05.Caption)
    
    '区分の表示を切り替える(今回引当分には「*」がつく)
    ' 在庫　　　　　 + ⇔   *
    ' 期限切れ在庫　 x ⇔ 切*
    ' 手入力の在庫　**
    Select Case Me.Controls("lbl区分_" & 行).Caption
    Case "**"   '別在庫はそのまま("**"：賞味期限、バッチ数の編集ができる)
    Case "*":   If Val(Me.Controls("lbl割当数_" & 行).Caption) = 0 Then Me.Controls("lbl区分_" & 行).Caption = "+"
    Case "+":   If Val(Me.Controls("lbl割当数_" & 行).Caption) <> 0 Then Me.Controls("lbl区分_" & 行).Caption = "*"
    Case "切*": If Val(Me.Controls("lbl割当数_" & 行).Caption) = 0 Then Me.Controls("lbl区分_" & 行).Caption = "x"
    Case "x":   If Val(Me.Controls("lbl割当数_" & 行).Caption) <> 0 Then Me.Controls("lbl区分_" & 行).Caption = "切*"
    End Select
    Me.Controls("lbl引当説明_" & 行).Caption = 引当マーク文言変換(Me.Controls("lbl区分_" & 行).Caption)

End Sub
Private Sub lbl割当数_01_Click()
    Call Input割当数("01")
End Sub
Private Sub lbl割当数_02_Click()
    Call Input割当数("02")
End Sub
Private Sub lbl割当数_03_Click()
    Call Input割当数("03")
End Sub
Private Sub lbl割当数_04_Click()
    Call Input割当数("04")
End Sub
Private Sub lbl割当数_05_Click()
    Call Input割当数("05")
End Sub

