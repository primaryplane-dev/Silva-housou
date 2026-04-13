Option Explicit

Private Sub UserForm_Initialize()
    cmdKeyBk.Caption = "Back" & vbCrLf & "Space"
    lblInputArea.Caption = P_InputTenKey
End Sub

'ボタン押下時の共通処理
Private Sub btnClick共通(ByVal i_BtnCaption As String)

    Select Case i_BtnCaption
    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "00"
        lblInputArea.Caption = lblInputArea.Caption & i_BtnCaption
    Case "←", "BackSpace"
        If lblInputArea.Caption <> "" Then
            lblInputArea.Caption = Mid(lblInputArea.Caption, 1, Len(lblInputArea.Caption) - 1)
        End If
    Case "C"
        lblInputArea.Caption = ""
    End Select

End Sub

Private Sub cmd決定_Click()
    P_InputTenKey = lblInputArea.Caption
    Unload Me
End Sub

Private Sub cmdキャンセル_Click()
    Unload Me
End Sub


'数字キー押下処理
Private Sub cmdKey0_Click()
    Call btnClick共通(ActiveControl.Caption)
End Sub

Private Sub cmdKey00_Click()
    Call btnClick共通(ActiveControl.Caption)
End Sub

Private Sub cmdKey1_Click()
    Call btnClick共通(ActiveControl.Caption)
End Sub

Private Sub cmdKey2_Click()
    Call btnClick共通(ActiveControl.Caption)
End Sub

Private Sub cmdKey3_Click()
    Call btnClick共通(ActiveControl.Caption)
End Sub

Private Sub cmdKey4_Click()
    Call btnClick共通(ActiveControl.Caption)
End Sub

Private Sub cmdKey5_Click()
    Call btnClick共通(ActiveControl.Caption)
End Sub

Private Sub cmdKey6_Click()
    Call btnClick共通(ActiveControl.Caption)
End Sub

Private Sub cmdKey7_Click()
    Call btnClick共通(ActiveControl.Caption)
End Sub

Private Sub cmdKey8_Click()
    Call btnClick共通(ActiveControl.Caption)
End Sub

Private Sub cmdKey9_Click()
    Call btnClick共通(ActiveControl.Caption)
End Sub

Private Sub cmdKeyBk_Click()
    Call btnClick共通("BackSpace")
End Sub

Private Sub cmdKeyClear_Click()
    Call btnClick共通("C")
End Sub
