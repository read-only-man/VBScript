' このコードは'nerstat -an'コマンドをほぼ同じ出力を生成する。
' なお、ターゲットマシンにSNMPとWMI SNMP Providerがインストールされている必要がある。
' ----- START CONFIGRATION -----
strComputerIP = "127.0.0.1"
' -----  END  CONFIGRATION -----
set objLocator = CreateObject("WbemScripting.SWbemLocator")
set objWMI = objLocator.ConnectServer("", "root/snmp/localhost")
set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
objNamedValueSet.Add "AgentAddress", strComputerIP
objNamedValueSet.Add "AgentReadCommunityName", "public"
objNamedValueSet.Add "AgentWriteCommunityName", "public"

WScript.Echo " Proto Local Address Foreign Address State"
set colTCPconns = objWMI.Instancesof("SNMP_RFC1213_MIB_tcpConnTable", , objNamedValueSet)
for each objConn in colTCPConns
  WScript.Echo " TCP " & objConn.tcpConnLocalAddress & ":" & objConn.tcpConnPort & " " & objConn.tcpConnRemAddress & ":" & _
  objConn.tcpConnRemPort & " " & objConn.tcpConnState
next
set colUDPconns = objWMI.Instancesof("SNMP_RFC1213_MIB_udpConnTable", , objNamedValueSet)
for each objConn in colUDPConns
  WScript.Echo " UDP " & objConn.udpConnLocalAddress & ":" & objConn.udpConnPort & " " & objConn.udpConnRemAddress & ":" & _
  objConn.udpConnRemPort & " " & objConn.udpConnState
next
