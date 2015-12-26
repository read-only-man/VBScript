' このコードはドメインユーザのロックを解除する。
' ----- START CONFIGRATION -----
strUseName = "<ユーザ名>"
strDomain = "<ドメインのNetBIOS名>"
' -----  END  CONFIGRATION -----
set objUser = GetObject("WinNT://" & strSomain & "/" & strUserName)
if objUser.IsAccountLocked = TRUE then
  objUser.IsAccountLocked = FALSE
  objUser.setInfo
  WScript.Echo "Account unlocked"
else
  WScript.Echo "Account is not locked"
end if
