' このコードは指定されたドメインにおいてFSMOロールを持つサーバ出力する。
' ----- START CONFIGRATION -----
strDomain = "<ドメインのDNS名>"
' -----  END  CONFIGRATION -----

set objRootDSE = GetObject("LDAP://" & strDomain & "/ROOTDSE")
strDomainDN = objRootDSE.Get("defaultNameingContext")
strSchemaDN = objRootDSE.Get("schemaNamingContext")
strConfigDN = objRootDSE.Get("configurationNamingContext")

' PDCエミュレータ
set objPDCFsmo = GetObject("LDAP://" & strDomainDN)
WScript.Echo "PDC Emulator: " & objPDCFsmo.fsmoroleowner

' RIDマスタ
set objRIDFsmo = GetObject("LDAP://cn=RID Manager$,cn=system," & strDomainDN)
WScript.Echo "RID Master: " & objRIDFsmo.fsmoroleowner

' スキーママスタ
set objSchemaFsmo = GetObject("LDAP://" & strSchemaDN)
WScript.Echo "Schema Master: " & objSchemaFsmo.fsmoroleowner

' インフラストラクチャマスタ
set objInfraFsmo = GetObject("LDAP://cn=Infrastructure," & strDomainDN)
WScript.Echo "Infrastructure Master: " &objInfraFsmo.fsmoroleowner

' ドメイン名前付けマスタ
set objDNFsmo = GetObject("LDAP://cn=Partitions," & strConfigDN)
WScript.Echo "Domain Naming Master:" & objDNFsmo.fsmoroleowner
