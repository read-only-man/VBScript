' ���̃R�[�h�͎w�肳�ꂽDNS�T�[�o�̓��v�����ׂĕ\������B
' ----- START CONFIGRATION -----
strServer = "<�T�[�o�[��>"
' -----  END  CONFIGRATION -----
set objDNS = GetObject("winmgmts:\\" & strServer & "\root\MicrosoftDNS")
set objDNSServer = objDNS.Get("MicrosoftDNS_Server.Name="".""")
set objStatus = objDNS.ExecQuery("Select * from MicrosoftDNS_Statistic ")
for each objStat in objStats
  WScript.Echo " " & objStat.Name & " : " & objStat.Value
next
