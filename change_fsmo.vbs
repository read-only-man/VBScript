' ���̃R�[�h�́APDC�G�~�����[�^�̃��[�����w�肳�ꂽ�h���C���R���g���[���֓]������
' ----- START CONFIGRATION -----
strNewOwner = "<�]����>"
' -----  END  CONFIGRATION -----
set objRootDSE = GetObject("LDAP://" & strNewOwner & "/ROOTDSE")
set myDomain = GetObject("LDAP://" & objRootDSE.get("defaultNamingContext"))
myDomainSid = myDomain.objectSid
objRootDSE.Put "becomePDC", myDomainSid
objRootDSE.SetInfo
