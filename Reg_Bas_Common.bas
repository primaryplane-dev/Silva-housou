Option Explicit

'Nullを空白に変換
Public Function NVL(ByVal i_変換前 As Variant) As String
    Dim str変換後       As String
    
    If Not IsNull(i_変換前) Then
        str変換後 = CStr(i_変換前)
    End If

    NVL = str変換後
End Function

'Nullをゼロに変換
Public Function NZ(ByVal i_変換前 As Variant) As Long
    Dim str変換後       As Long
    
    If Not IsNull(i_変換前) Then
        str変換後 = CLng(i_変換前)
    End If

    NZ = str変換後
End Function

Public Function 日付変換(ByVal i_変換前 As Variant) As Date
    日付変換 = 0

    If IsNull(i_変換前) Then Exit Function
    If Val(i_変換前) = 0 Then Exit Function
    
    日付変換 = CDate(Format(i_変換前, "0000/00/00"))

End Function

Public Function 引当マーク文言変換(ByVal i_変換前 As String) As String
    Dim str変換後       As String
    
    Select Case i_変換前
    Case "*":   str変換後 = "自動引当"
    Case "**":  str変換後 = "手動引当"
    Case "x":   str変換後 = "出荷期限切れ在庫"
    Case "切*": str変換後 = "出荷期限切れ在庫を出荷"
    Case "+":   str変換後 = ""
    End Select

    引当マーク文言変換 = str変換後
End Function

'ボタン共通処理
Public Sub M_sbClickButton(ByVal strFileName As String)
    Dim FSO                 As Object
    Dim App                 As Object
    Dim strFullPath         As String           'ファイルのフルパス
    
    Application.Cursor = xlWait
    Set FSO = CreateObject("Scripting.FileSystemObject")

'   Call Write設定ini                           Iniを使う場合はここで出力する

'    起動するファイルのフルパス (カレントディレクトリ)
    strFullPath = FSO.BuildPath(Application.ThisWorkbook.Path, strFileName)

'    二重起動対応                               '2017/05/11 Add 榊原
    On Error GoTo ErrHandler

'    サーバから最新ファイルをコピーする
    Call FSO.CopyFile(PB_SERVER & strFileName, strFullPath, True)

    On Error GoTo 0
    
'    ログ出力
    Call SubLogging(strFileName)

'    EXCELを起動する
    Set App = CreateObject("Excel.Application")
    App.Visible = True
    App.Workbooks.Open strFullPath

    Set FSO = Nothing
    Application.Cursor = xlDefault
    Exit Sub

'    二重起動対応                               '2017/05/11 Add 榊原
ErrHandler:
    If Err.Number = 70 Then                     '書き込み不可
        MsgBox strFileName & vbCrLf & "は起動しています", vbExclamation
        Exit Sub
    Else
        MsgBox strFileName & vbCrLf & "は起動できませんでした" & vbCrLf & " (" & Err.Number & " )", vbExclamation
    End If
End Sub
