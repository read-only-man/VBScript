' このコードは、PDCエミュレータのロールを指定されたドメインコントローラへ転送する
' ----- START CONFIGRATION -----
strNewOwner = "<転送先>"
' -----  END  CONFIGRATION -----
set objRootDSE = GetObject("LDAP://" & strNewOwner & "/ROOTDSE")
set myDomain = GetObject("LDAP://" & objRootDSE.get("defaultNamingContext"))
myDomainSid = myDomain.objectSid
objRootDSE.Put "becomePDC", myDomainSid
objRootDSE.SetInfo
