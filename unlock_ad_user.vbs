' ���̃R�[�h�̓h���C�����[�U�̃��b�N����������B
' ----- START CONFIGRATION -----
strUseName = "<���[�U��>"
strDomain = "<�h���C����NetBIOS��>"
' -----  END  CONFIGRATION -----
set objUser = GetObject("WinNT://" & strSomain & "/" & strUserName)
if objUser.IsAccountLocked = TRUE then
  objUser.IsAccountLocked = FALSE
  objUser.setInfo
  WScript.Echo "Account unlocked"
else
  WScript.Echo "Account is not locked"
end if
