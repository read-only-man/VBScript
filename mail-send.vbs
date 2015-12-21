' メール送信用モジュール
' このスクリプトを実行するにはCDO（Collaboration Data Objects）がクライアントコンピュータにインストールされている必要がある。
' そのためには Outlookをインストールするか、MicroSoftのサイトからCDOを個別にインストールする必要がある。

' ----- START CONFIGRATION -----
strFrom ="test@mail.com"
strTo   = "someone@mail.com"
strSub  = "mail title"
strBody = "this is mail test" & VBCRLF & "Hello"
strSMTPServer = "smtp.server.com"
' ----- END CONFIGRATION -------

set objEmail = CreateObject"CDO.Message")
objEmail.From = strFrom
objEmail.To = strTo
objEmail.Subject = strSub
objEmail.TextBody = strBody
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer
objEmail.Configuration.Fields.Update
objEmail.Send
WScript.Echo "Email sent"

objEmail = Nothing
