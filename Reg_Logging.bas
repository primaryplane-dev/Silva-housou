Option Explicit

'*****ログ出力
Public Sub SubLogging(ByVal sLoggingMessage As String)
    If sLoggingMessage = vbNullString Then sLoggingMessage = "メッセージ無し"
    Call SubWriteLog(sLoggingMessage)
End Sub

'*****メッセージをログファイルに出力
Sub SubWriteLog(ByVal sLogMsg As String)
    Dim FsObject As Object
    Dim FsLOG As Object
    Dim sLogFile As String
    Dim sHostName As String
    Dim sUserName As String
    Dim sPath As String
    
    sPath = PB_SERVER
    sHostName = Environ("COMPUTERNAME")
    sUserName = Environ("USERNAME")
    
    '***ログファイル名生成
    sLogFile = sPath & sHostName & "_AccessLog.log"
    
    Set FsObject = CreateObject("Scripting.FileSystemObject")
    
    ''***ログファイルがなければ作成
    If FsObject.FileExists(sLogFile) = False Then
        FsObject.CreateTextFile sLogFile
    End If
    
    ''***ログファイルを追記で開く
    Set FsLOG = FsObject.OpenTextFile(sLogFile, 8)  'ForAppending

    ''***PC名＋IPアドレス＋ユーザー名＋日時＋メッセージを書き込む
    FsLOG.WriteLine sHostName & Chr(44) & _
                    FncGetIPaddress & Chr(44) & _
                    sUserName & Chr(44) & _
                    Now & Chr(44) & _
                    sLogMsg
    
    ''***ログファイルを閉じる
    FsLOG.Close
      
    ''***後処理
    Set FsLOG = Nothing
    Set FsObject = Nothing
End Sub

'*****DNSサーバーを指定してIPアドレスを取得する
Function FncGetIPaddress() As String
  Dim strDNSServer0 As String
  Dim strIp As String
  Dim strChk As String
  Dim objNic As Object
  Dim oneNic As Object
  
  FncGetIPaddress = ""
  
  strDNSServer0 = "10.2.1.12" '←優先DNSサーバーがコレのIPを調べる
  strIp = "該当無し"

  Set objNic = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where (IPEnabled = TRUE)")
  For Each oneNic In objNic
    'DefaultGatewayの設定が無いNICは無視
    If IsError(oneNic.DNSServerSearchOrder(0)) = False Then
      strChk = oneNic.DNSServerSearchOrder(0)
      If InStr(strChk, strDNSServer0) > 0 Then
        strIp = oneNic.ipaddress(0)
        Exit For
      End If
    End If
  Next
  FncGetIPaddress = "IPアドレス：" & strIp

End Function


