' このコードはローカルDNSキャッシュをフラッシュする。
' これを目的としたスクリプティングインタフェイスはないため、
' シェルを起動してipconfig /flushdnsコマンドを呼び出している。
strCommand = "ipconfig /flushdns"
run_command(strCommand)

' このコードはローカルホストのDNSレコードを登録する
' これを目的としたスクリプティングインタフェイスはないため、
' シェルを起動してipconfi /registerdnsコマンドを呼び出している。
strCommand = "ipconfig /registerdns"
run_command(strCommand)

function run_command(strCommand)
  set objWshShell = WScript.CreateObject("WScript.Shell")
  intRC= objWshShell.Run(strCommand, 0 ,TRUE)
  if intRC <> 0 then
    WScript.Echo "Error returned from running the command <" & strCommand & "> : " & intRC
  else
    WScript.Echo "Command excuted sucessfully <" & strCommand & ">"
  end if
end function
