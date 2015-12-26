' 全接続のネットワーク構成情報を取得する。
' ----- START CONFiGRATION -----
strComputer = "."
' -----  END  CONFIGRATION -----
' Windows Management Instrumentation（WMI）は、Windows Driver Modelへの拡張の一種で、
' システムの構成要素について情報収集と通知を行うオペレーティングシステムのインタフェースを提供する。
set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
set colNAs = objWMI.InstanceOf("Win32_NetworkAdapter")
for each objNA in colNAs
  Wscript.Echo ObjNa.Name
  Wscript.Echo  " Description: " & objNA.Description
  Wscript.Echo  " Product Name: " & objNA.ProductName
  Wscript.Echo  " Manufacture: " & objNA.Manufacture
  Wscript.Echo  " Adapter Type: " & objNA.AdapterType
  Wscript.Echo  " AutoSense:" & objNA.AutoSense
  Wscript.Echo  " MAC Address: " & objNA.MACAddress
  Wscript.Echo  " Maximum Speed: " & objNA.MaxSpeed
  Wscript.Echo  " Conn Status: " & objNA.NetConnectionStatus
  Wscript.Echo  " Service Name: " & objNA.ServiceName
  Wscript.Echo  " Speed: " & objNA.Speed
  
  set colNACs = objWMI.ExecQuery(" select * from " &_
    " Win32_NetworkAdapterConfigration " &_
    " where Index = " & objNA.Index)

  ' colNACの項目は１つだけのはず
  for each objNAC in colNACs
    if IsArray(objNAC.IPAddress) then
      for each strAddress in objNAC.IPAddress
        Wscript.Echo " Network Addr: " & strAddress
      next
    end if
    Wscript.Echo  " IP Metric: " & objNA.IPConnectionMetric
    Wscript.Echo  " IP Enabled: " & objNA.IPEnabled
    Wscript.Echo  " Filter: " & objNA.IPFilterSecurityEnabled
    Wscript.Echo  " Port Security: " & objNA.IPPortSecurityEnabled
    if IsArray(objNAC.IPSubnet) then
      for each strSubnet in objNAC.IPSubnet
        Wscript.Echo " Subnet Mask: " & strSubnet
      next
    end if
    if IsArray(objNAC.DefaultIPGateway) then
      for each strGW in objNAC.DefaultIPGateway
        Wscript.Echo " Gateway Addr: " & strGW
      next
    end if
    Wscript.Echo  " Database Path: " & objNA.DatabasePath
    Wscript.Echo  " DHCP Enabled: " & objNA.DHCPEnabled
    Wscript.Echo  " Lease Expires: " & objNA.DHCPLeaseExpires
    Wscript.Echo  " Lease Obtained: " & objNA.DHCPLeasePbtained
    Wscript.Echo  " DHCP Server: " & objNA.DHCPServer
    Wscript.Echo  " DNA Domain: " & objNA.DNSDomain
    Wscript.Echo  " DNS For WINS: " & objNA.DNSEnabledForWINSResolution
    Wscript.Echo  " DNS Host Name: " & objNA.DNSHostName
    if IsArray(objNAC.DNSDomainSuffixSearchOrder) then
      for each strName in objNAC.DNSDomainSuffixSearchOrder
        Wscript.Echo " DNS Suffix Search Order: " & strName
      next
    end if
    if IsArray(objNAC.DNSServerSearchOrder) then
      for each strName in objNAC.DNSServerSearchOrder
        Wscript.Echo " DNS Server Search Order: " & strName
      next
    end if
    Wscript.Echo  " Domain DNS Reg Enabled: " & objNA.DomainDNSRegistrationEnabled
    Wscript.Echo  " Full DNS Req Enabled: " & objNA.FullDNSRegistrationEnabled
    Wscript.Echo  " LMHosts Lookup: " & objNA.WINSHostLookupFile
    Wscript.Echo  " WINS Lookup File: " & objNA.WINSHostLookupFile
    Wscript.Echo  " WINS Scope ID: " & objNAWINSScopeID
    Wscript.Echo  " WINS Primary Server: " & objNA.WINSPrimaryServer
    Wscript.Echo  " WINS Secondary: " & objNA.WINSSecondaryServer
  next
next