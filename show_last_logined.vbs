' ���̃R�[�h�̓��[�U�̍ŏI���O�I��������\������B
' ----- START CONFIGRATION -----
strUserDN = "<���[�UDN>"
' -----  END  CONFIGRATION -----

setObjUser = GetObject("LDAP://" & strUserDN)
set objLogon = objUser.Get("lastLogonTimestamp")
intLogonTime = objLogon.HighPart * (2^32) + objLogon.LowPart
intLogonTime = intLogonTime / (60 * 10000000)
intLogonTime = intLogonTime / 1440
WScript.Echo "Approx last logon timestamp: " & intLogonTime + #1/1/1601#
