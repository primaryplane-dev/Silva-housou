Option Explicit

Private Sub Workbook_Open()
    
    ThisWorkbook.Activate
    
    Application.EnableEvents = False
   
    '変数クリア
    P_終了ボタン押下 = False
    
    'リストクリア
    Call 一覧クリア
    
    ReDim P_製品(0)

    Application.EnableEvents = True
    
    'シート選択
    st01List.Select
    Cells(1, 1).Select

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    If Not (Dic運送会社 Is Nothing) Then Set Dic運送会社 = Nothing

    If P_終了ボタン押下 = False Then
        MsgBox "閉じるボタンは使用できません。" & vbCrLf & "リストシートの終了ボタンで終了してください。", vbCritical
        Cancel = True
    End If

End Sub



