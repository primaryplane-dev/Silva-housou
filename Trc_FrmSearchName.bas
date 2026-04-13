Option Explicit

'抽出した名称を保存しておく
Private Dic選択                 As Object
Private ChangeFLG               As Boolean          'LstNameの複数選択の制御に使う
Private m_SearchKey             As String

Private Sub 選択リスト作成(ByVal i_品名 As String)
    Dim CN                  As Object
    Dim RS                  As Object
    Dim strSQL              As String
    Dim i                   As Integer
    Dim STM                 As Object
    
    Set CN = CreateObject("ADODB.Connection")
    Set RS = CreateObject("ADODB.Recordset")
    CN.CursorLocation = adUseClient
    CN.Open P_接続文字列

    strSQL = ""
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "   RHHNO AS HNO"
    strSQL = strSQL & "  ,RHHNM AS HNM"
    strSQL = strSQL & " FROM " & P_LIB & ".SRHP01 "
    strSQL = strSQL & " WHERE RHDLT = ''"
    strSQL = strSQL & "   AND RHHNM LIKE '%" & i_品名 & "%'"
    strSQL = strSQL & " ORDER BY 2"
    Debug.Print strSQL
    Set RS = New ADODB.Recordset
    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    If RS.RecordCount > 0 Then
        i = -1
        lstName.Clear
        Do While Not RS.EOF
            i = i + 1
            With lstName
                .AddItem
                .List(i, 0) = RS("HNO")
                .List(i, 1) = RTrim(RS("HNM"))
            End With
            RS.MoveNext
        Loop
    End If
    RS.Close:    Set RS = Nothing
    CN.Close:    Set CN = Nothing
    
End Sub

Private Sub cmd検索_Click()
    Call 選択リスト作成(m_SearchKey)
End Sub

'-------------------------------------------------------------------------------
'　データ選択ボタン押下
'-------------------------------------------------------------------------------
Private Sub cmdデータ選択_Click()
    Dim i       As Integer
    Dim j       As Integer

    '選択された行を保存する
    ReDim P_製品(lstName.ListCount)
    With lstName
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                j = j + 1
                P_製品(j).CD = .List(i, 0)
                P_製品(j).NM = .List(i, 1)
            End If
        Next i
        '使わなかったエリアを消す
        ReDim Preserve P_製品(j)
    End With

    '入力チェック
    If UBound(P_製品) = 0 Then MsgBox "製品を選択してください": Exit Sub
    If UBound(P_製品) > 10 Then MsgBox "10個まで！":            Exit Sub

    P_検索FLG = True
    Unload Me
End Sub

'-------------------------------------------------------------------------------
'　キャンセルボタン押下
'-------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'-------------------------------------------------------------------------------
' かなボタン押下
'-------------------------------------------------------------------------------
Sub SetSearchKey(strKEYName As String)
    m_SearchKey = m_SearchKey & strKEYName
    txtSearchKey.Text = m_SearchKey
End Sub

Private Sub btnA_Click()
    Call SetSearchKey("ｱ")
End Sub

Private Sub btnI_Click()
    Call SetSearchKey("ｲ")
End Sub

Private Sub btnU_Click()
    Call SetSearchKey("ｳ")
End Sub

Private Sub btnE_Click()
    Call SetSearchKey("ｴ")
End Sub

Private Sub btnO_Click()
    Call SetSearchKey("ｵ")
End Sub

Private Sub btnKA_Click()
    Call SetSearchKey("ｶ")
End Sub

Private Sub btnKI_Click()
    Call SetSearchKey("ｷ")
End Sub

Private Sub btnKU_Click()
    Call SetSearchKey("ｸ")
End Sub

Private Sub btnKE_Click()
    Call SetSearchKey("ｹ")
End Sub

Private Sub btnKO_Click()
    Call SetSearchKey("ｺ")
End Sub

Private Sub btnSA_Click()
    Call SetSearchKey("ｻ")
End Sub

Private Sub btnSI_Click()
    Call SetSearchKey("ｼ")
End Sub

Private Sub btnSU_Click()
    Call SetSearchKey("ｽ")
End Sub

Private Sub btnSE_Click()
    Call SetSearchKey("ｾ")
End Sub

Private Sub btnSO_Click()
    Call SetSearchKey("ｿ")
End Sub

Private Sub btnTA_Click()
    Call SetSearchKey("ﾀ")
End Sub

Private Sub btnTI_Click()
    Call SetSearchKey("ﾁ")
End Sub

Private Sub btnTU_Click()
    Call SetSearchKey("ﾂ")
End Sub

Private Sub btnTE_Click()
    Call SetSearchKey("ﾃ")
End Sub

Private Sub btnTO_Click()
    Call SetSearchKey("ﾄ")
End Sub

Private Sub btnNA_Click()
    Call SetSearchKey("ﾅ")
End Sub

Private Sub btnNI_Click()
    Call SetSearchKey("ﾆ")
End Sub

Private Sub btnNU_Click()
    Call SetSearchKey("ﾇ")
End Sub

Private Sub btnNE_Click()
    Call SetSearchKey("ﾈ")
End Sub

Private Sub btnNO_Click()
    Call SetSearchKey("ﾉ")
End Sub

Private Sub btnHA_Click()
    Call SetSearchKey("ﾊ")
End Sub

Private Sub btnHI_Click()
    Call SetSearchKey("ﾋ")
End Sub

Private Sub btnFU_Click()
    Call SetSearchKey("ﾌ")
End Sub

Private Sub btnHE_Click()
    Call SetSearchKey("ﾍ")
End Sub

Private Sub btnHO_Click()
    Call SetSearchKey("ﾎ")
End Sub

Private Sub btnMA_Click()
    Call SetSearchKey("ﾏ")
End Sub

Private Sub btnMI_Click()
    Call SetSearchKey("ﾐ")
End Sub

Private Sub btnMU_Click()
    Call SetSearchKey("ﾑ")
End Sub

Private Sub btnME_Click()
    Call SetSearchKey("ﾒ")
End Sub

Private Sub btnMO_Click()
    Call SetSearchKey("ﾓ")
End Sub

Private Sub btnYA_Click()
    Call SetSearchKey("ﾔ")
End Sub

Private Sub btnYU_Click()
    Call SetSearchKey("ﾕ")
End Sub

Private Sub btnYO_Click()
    Call SetSearchKey("ﾖ")
End Sub

Private Sub btnRA_Click()
    Call SetSearchKey("ﾗ")
End Sub

Private Sub btnRI_Click()
    Call SetSearchKey("ﾘ")
End Sub

Private Sub btnRU_Click()
    Call SetSearchKey("ﾙ")
End Sub

Private Sub btnRE_Click()
    Call SetSearchKey("ﾚ")
End Sub

Private Sub btnRO_Click()
    Call SetSearchKey("ﾛ")
End Sub

Private Sub btnWA_Click()
    Call SetSearchKey("ﾜ")
End Sub

Private Sub btnWO_Click()
    Call SetSearchKey("ｦ")
End Sub

Private Sub btnNn_Click()
    Call SetSearchKey("ﾝ")
End Sub

Private Sub btnLI_Click()
    Call SetSearchKey("ｨ")
End Sub

Private Sub btnLE_Click()
    Call SetSearchKey("ｪ")
End Sub

Private Sub btnLTU_Click()
    Call SetSearchKey("ｯ")
End Sub

Private Sub btnLYA_Click()
    Call SetSearchKey("ｬ")
End Sub

Private Sub btnLYU_Click()
    Call SetSearchKey("ｭ")
End Sub

Private Sub btnLYO_Click()
    Call SetSearchKey("ｮ")
End Sub

Private Sub btnDa_Click()
    Call SetSearchKey("ﾞ")
End Sub

Private Sub btnPa_Click()
    Call SetSearchKey("ﾟ")
End Sub

Private Sub btnCho_Click()
    'Call SetSearchKey("ー")
    Call SetSearchKey("ｰ")
End Sub

'クリア
Private Sub btnClear_Click()
    m_SearchKey = ""
    txtSearchKey.Text = m_SearchKey
End Sub

'BackSpace
Private Sub btnBackSpace_Click()
    '末尾１文字を除去する
    If m_SearchKey <> "" Then
        m_SearchKey = Mid(m_SearchKey, 1, Len(m_SearchKey) - 1)
    End If
    txtSearchKey.Text = m_SearchKey
End Sub

'-------------------------------------------------------------------------------
'リストの複数選択対応（タッチパネルではctrlキーが使えないため)
'-------------------------------------------------------------------------------
Private Sub lstName_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ChangeFLG = True
End Sub

Private Sub lstName_Change()
    If ChangeFLG = False Then Exit Sub

    ChangeFLG = False
    Call Lst選択行の記憶(lstName.ListIndex)
    Call Lst選択行の設定
End Sub

Private Sub Lst選択行の記憶(ByVal idx As Integer)
    Dim strKey      As String
    If Dic選択 Is Nothing Then Set Dic選択 = CreateObject("Scripting.Dictionary")
    
    strKey = CStr(idx)
    If Not Dic選択.exists(strKey) Then
        Dic選択.Add strKey, "1"             '値はなんでも良い
    Else
        Dic選択.Remove strKey
    End If
End Sub

Private Sub Lst選択行の設定()
    Dim i As Integer
    For i = 0 To lstName.ListCount - 1
        If Dic選択.exists(CStr(i)) Then
            lstName.Selected(i) = True
        Else
            lstName.Selected(i) = False
        End If
    Next
End Sub

