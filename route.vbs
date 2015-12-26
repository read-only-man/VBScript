' このコードは"route print"コマンドと同様の情報を出力する。
' Win32_IP4RouteTableクラスはWindows Server 2003で新たに追加されたものなので、
' このスクリプトはWindows Server 2000では動作しない。
' ----- START CONFIGRATION -----
strComputer = "."
' -----  END  CONFIGRATION -----
set objWMI = GerObject("winmgmts:\\" & strComputer & "\root\civmv2")
set colRoutes = objWMI.InstancesOf("Win32_IP4RouteTable")
for each objTRoutes in colRoutes
  set colNetworkAdapters = objWMI.ExecQuery("select * from Win32_NetworkAdapterConfigration " &_
    "where Interfaceindex = " & objRoute.InterfaceIndex )
  for each objNetworkAdapter un colNetworkAdapters
    for each strIP in objNetworlAdapter.IPAddress
      WScript.Echo "Interface: " & strIP
    next
  next

  WScript.Echo "Network: " & objRoute.Destination
  WScript.Echo "NetMask: " & objRoute.Mask
  WScript.Echo "Gateway: " & objRoute.NextHop
  WScript.Echo "Metric: "  & objRoute.Metric1

  WScript.Echo "Age: "             & objRoute.Age
  WScript.Echo "Description: "     & objRoute.Description
  WScript.Echo "Information: "     & objRoute.Information
  WScript.Echo "Interface Index: " & objRoute.InterfaceIndex

  WScript.Echo "Metric2: "  & objRoute.Metric2
  WScript.Echo "Metric3: "  & objRoute.Metric3
  WScript.Echo "Metric4: "  & objRoute.Metric4
  WScript.Echo "Metric5: "  & objRoute.Metric5

  WScript.Echo "Name: "      & objRoute.Name
  WScript.Echo "Protocol": " & objRoute.Protocol
  WScript.Echo "Status: "    & objRoute.Status
  WScript.Echo "Type: "      & objRoute.Type
next