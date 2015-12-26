' ���̃R�[�h�͎w�肳�ꂽ�h���C���ɂ�����FSMO���[�������T�[�o�o�͂���B
' ----- START CONFIGRATION -----
strDomain = "<�h���C����DNS��>"
' -----  END  CONFIGRATION -----

set objRootDSE = GetObject("LDAP://" & strDomain & "/ROOTDSE")
strDomainDN = objRootDSE.Get("defaultNameingContext")
strSchemaDN = objRootDSE.Get("schemaNamingContext")
strConfigDN = objRootDSE.Get("configurationNamingContext")

' PDC�G�~�����[�^
set objPDCFsmo = GetObject("LDAP://" & strDomainDN)
WScript.Echo "PDC Emulator: " & objPDCFsmo.fsmoroleowner

' RID�}�X�^
set objRIDFsmo = GetObject("LDAP://cn=RID Manager$,cn=system," & strDomainDN)
WScript.Echo "RID Master: " & objRIDFsmo.fsmoroleowner

' �X�L�[�}�}�X�^
set objSchemaFsmo = GetObject("LDAP://" & strSchemaDN)
WScript.Echo "Schema Master: " & objSchemaFsmo.fsmoroleowner

' �C���t���X�g���N�`���}�X�^
set objInfraFsmo = GetObject("LDAP://cn=Infrastructure," & strDomainDN)
WScript.Echo "Infrastructure Master: " &objInfraFsmo.fsmoroleowner

' �h���C�����O�t���}�X�^
set objDNFsmo = GetObject("LDAP://cn=Partitions," & strConfigDN)
WScript.Echo "Domain Naming Master:" & objDNFsmo.fsmoroleowner
