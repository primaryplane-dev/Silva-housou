Option Explicit

Public Function DB更新(ByRef 出荷Rec As 出荷Record, ByRef CN As Object) As Boolean
    Dim strSQL              As String
    Dim 行                  As Long
    Dim 処理件数            As Long
    Dim strMSG              As String

    DB更新 = False

    Err.Clear
    On Error Resume Next

    '■変更箇所をデータベースに反映する
    ' 11列目（車両積荷前衛生点検:1/0）、12列目（逸脱事項:フリー入力）はAS連携で直接InsertSQL/UpdateSQLに渡す

    'LIBSMF17.SZSP01:出荷データ
    strSQL = InsertSQL(出荷Rec)
    CN.Execute strSQL, 処理件数, &H80
    If Err.Number <> 0 Then GoTo DB更新_Err
    If 処理件数 = 0 Then GoTo DB更新_Err

    'LIBSMF17.SSZP01:在庫データ
    strSQL = UpdateSQL_在庫(出荷Rec)
    CN.Execute strSQL, 処理件数, &H80
    If Err.Number <> 0 Then GoTo DB更新_Err
    If 処理件数 = 0 Then
        strSQL = InsertSQL_在庫(出荷Rec)
        CN.Execute strSQL, 処理件数, &H80
        If Err.Number <> 0 Then GoTo DB更新_Err
        If 処理件数 = 0 Then GoTo DB更新_Err
    End If

    '終わり
    DB更新 = True
    GoTo DB更新_Exit


DB更新_Err:
    strMSG = strSQL & Err.Description
    Debug.Print strMSG
    Err.Clear
    GoTo DB更新_Exit

DB更新_Exit:
    If strMSG <> "" Then Call MsgBox(strMSG, vbOKOnly)
End Function

Public Function DB更新_運送会社(ByRef 出荷Rec As 出荷Record, ByRef CN As Object) As Boolean
    Dim strSQL              As String
    Dim 処理件数            As Long
    Dim strMSG              As String

    DB更新_運送会社 = False

    Err.Clear
    On Error Resume Next

    '■運送会社を全行まとめて更新する
    'LIBSMF17.SZSP01:出荷データ
    strSQL = UpdateSQL(出荷Rec)
    CN.Execute strSQL, 処理件数, &H80
    If Err.Number <> 0 Then GoTo DB更新_運送会社_Err
    
    '終わり
    DB更新_運送会社 = True
    GoTo DB更新_運送会社_Exit

DB更新_運送会社_Err:
    strMSG = strSQL & Err.Description
    Debug.Print strMSG
    Err.Clear
    GoTo DB更新_運送会社_Exit

DB更新_運送会社_Exit:
    If strMSG <> "" Then Call MsgBox(strMSG, vbOKOnly)
End Function

