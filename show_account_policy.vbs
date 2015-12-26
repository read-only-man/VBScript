' ���̃R�[�h�́A�A�J�E���g���b�N�A�E�g�|���V�[�ƃp�X���[�h�|���V�[�̌��ݐݒ�l��\������B
' ----- START CONFIGRATION -----
strDomain = "<�h���C����DN>"
' -----  END  CONFIGRATION -----
set objRootDSE = GetObject("LDAP://" & strDomain & "/RootDSE")
set objDomain = GetObject("LDAP://" & objRootDSE.Get("defaultNamingContext"))

' �h���C���̃p�X���[�h�|���V�[�����ƃ��b�N�A�E�g�|���V�[�������L�[�A
' ���̒P�ʁi���Ȃǁj��l�Ɏ��n�b�V��
set objDomAttrHash = CreateObject("Scripting.Dictionary")
objDomAttrHash.Add "lockoutDuration", "minutes"
objDomAttrHash.Add "lockoutThreshold", "attempts"
objDomAttrHash.Add "lockoutObservationWindow", "minutes"
objDomAttrHash.Add "maxPwdAge", "minutes"
objDomAttrHash.Add "minPwdAge", "minutes"
objDomAttrHash.Add "minPwdLength, "characters"
objDomAttrHash.Add "pdwHistoryLength", "remembered"
objDomAttrHash.Add "pwdProperties", " "

' �e�����Ƀ��[�v�������A�o�͂���B
for each strAttr in objDomAttrHash.Keys
  if IsObject(objDomain.Get(strAttr)) then
    set ObjLargeInt = objDomain.Get(strAttr)
    if objLargeInt.Lowpart = 0 then
      value = 0
    else
      value = Abs(objLargeInt.HighPart * 2^32 + objLargeInt.LowPart)
      value = int(value / 10000000)
      value = int(value /60)
    end if
  else
    value = objDomain.Get(strAttr)
  end if
  WScript.Echo strAttr & " = " & value & " " & objDomAttrHash(strAttr)
next

' DOMAIN_PASSWORD_INFORMATION�Ɋ�Â��萔
set objDomainPassHash = CreateObject("Scripting.Dicionary")
objDomainPassHash.Add "DOMAIN_PASSWORD_COMPLEX", &h1
objDomainPassHash.Add "DOMAIN_PASSWORD_NO_ANON_CHANGE", &h2
objDomainPassHash.Add "DOMAIN_PASSWORD_NO_CLEAR_CHANGE", &h4
objDomainPassHash.Add "DOMAIN_PASSWORD_LOCKOUT_ADMINS", &h8
objDomainPassHash.Add "DOMAIN_PASSWORD_STORE_CLEARTEXT", &h16
objDomainPassHash.Add "DOMAIN_REFUSE_PASSWORD_CHANGE", &h32

' PwdProperties�����͕����̐ݒ��ێ�����t���O�Ȃ̂ŁA���ʂɏ�������B
for each strFlag IN objDomPassHash.Keys
  if objDomPassHash(strFlag) and objDomain.Get("PwdProperties") then
    WScript.Echo " " & strFlag & " is enabled"
  else
    WScript.Echo " " & strFlag & " is disabled"
  end if
next