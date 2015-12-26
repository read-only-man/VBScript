' このコードは、アカウントロックアウトポリシーとパスワードポリシーの現在設定値を表示する。
' ----- START CONFIGRATION -----
strDomain = "<ドメインのDN>"
' -----  END  CONFIGRATION -----
set objRootDSE = GetObject("LDAP://" & strDomain & "/RootDSE")
set objDomain = GetObject("LDAP://" & objRootDSE.Get("defaultNamingContext"))

' ドメインのパスワードポリシー属性とロックアウトポリシー属性をキー、
' その単位（分など）を値に持つハッシュ
set objDomAttrHash = CreateObject("Scripting.Dictionary")
objDomAttrHash.Add "lockoutDuration", "minutes"
objDomAttrHash.Add "lockoutThreshold", "attempts"
objDomAttrHash.Add "lockoutObservationWindow", "minutes"
objDomAttrHash.Add "maxPwdAge", "minutes"
objDomAttrHash.Add "minPwdAge", "minutes"
objDomAttrHash.Add "minPwdLength, "characters"
objDomAttrHash.Add "pdwHistoryLength", "remembered"
objDomAttrHash.Add "pwdProperties", " "

' 各属性にループをかけ、出力する。
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

' DOMAIN_PASSWORD_INFORMATIONに基づく定数
set objDomainPassHash = CreateObject("Scripting.Dicionary")
objDomainPassHash.Add "DOMAIN_PASSWORD_COMPLEX", &h1
objDomainPassHash.Add "DOMAIN_PASSWORD_NO_ANON_CHANGE", &h2
objDomainPassHash.Add "DOMAIN_PASSWORD_NO_CLEAR_CHANGE", &h4
objDomainPassHash.Add "DOMAIN_PASSWORD_LOCKOUT_ADMINS", &h8
objDomainPassHash.Add "DOMAIN_PASSWORD_STORE_CLEARTEXT", &h16
objDomainPassHash.Add "DOMAIN_REFUSE_PASSWORD_CHANGE", &h32

' PwdProperties属性は複数の設定を保持するフラグなので、特別に処理する。
for each strFlag IN objDomPassHash.Keys
  if objDomPassHash(strFlag) and objDomain.Get("PwdProperties") then
    WScript.Echo " " & strFlag & " is enabled"
  else
    WScript.Echo " " & strFlag & " is disabled"
  end if
next