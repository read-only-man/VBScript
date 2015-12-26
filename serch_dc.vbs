' このコードは、ドメインにおいて、このスクリプトが実行されているコンピュータから
' 最も近いドメインコントローラを検索する。
' ----- START CONFIGRATION -----
strDomain = "<ドメインのDNS名>"
' -----  END  CONFIGRATION -----
set objIadsTools = CreateObject("IADsTools.DCFunctions")
objIadsTools.DsGetDcName(Cstr(strDomain))
WScript.Echo "DC: " & objIadsTools.DcName
WScript.Echo "DC Site: " & objIadsTools.DcSiteName
WScript.Echo "Client Site: " & objIadsTools.ClientSiteName
