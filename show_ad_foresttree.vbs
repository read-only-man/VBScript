' このコードはフォレストツリーの階層情報を出力する

' BEGIN SECTION 1
set objRootDSE = GetObject("LDAP://RootDSE")
strBase = "<LDAP://cn=Partitions" & objRootDSE.Get("ConfigurationNamingContext") & ">;"
strFilter = "(&(objectcategory=crossRef)(systeFlags=3));"
strAttrs = "name,trustParent,nCName,dnsRoot,distinguishedName;"
strScope = "onelevel"
set objConn = CreateObject("ADODB.Connection")
objConn.Provider = "ADsDSOObject"
objConn.Open "Active Directory Provider"
set objRS = objConn.Execute(strBase & strFilter & strAttrs & strScope)
objRS.MoveFirst
'  END  SECTION 1

' BEGIN SECTION 2
set dicSubDomainTrue = CreateObject("Scripting.Dictionary")
set dicDomainHierarchy = CreateObject("Scripting.Dictionary")
set dicDomainRoot = CreateObject("Scripting.Dictionary")
'  END  SECTION 2

' BEGIN SECTION 3
while not objRS.EOF
  dicDomainRoot.Add objRS.Fields("name").Value, objRS.Fields("nCname").Value
  if objRS.Fields("trustParent").Value <> "" then
    dicSubDomainTrue,Add objRS.Fields("name").Value, 0
    set objDomainParent = GetObject("LDAP://" & objRS.Fields("trustParent").Value)
    dicDomainHierarchy.Add objRS.Fields("name").Value, objDomainParent.Get("name")
  else
    dicSubDomainTrue.Add objRS.Fields("name").Value, 1
  end if
  objRS.MoveNext
wend
'  END  SECTION 3

' BEGIN SECTION 4
for each strDomain in dicSubDomainTrue
  if dicSubDomainTrue(strDomain) = 1 then
    DisplayDomains strDomain, "", dicDomainHierarchy, dicDomainRoot
  end if
next
'  END  SECTION 4

' DisplayDomains
Function DisplayDomains ( strDomain, strSpaces, dicDomainHierarchy, dicDomainRoot)
  WScript.Echo strSpaces & strDomain
  DisplayObjects "LDAP://" & dicDomainRoot(strDomain), " " & strSpaces
  for each strD in dicDomainHierarchy
    if dicDomainHierarchey(strD) = strDomain then
      Displaydomains objChildObject.ADsPath, strSpaces & " ", dicDomainHierarchy, dicDomainRoot
    end if
  next
end Function

' DisplayObjects関数には、子オブジェクトを表示するオブジェクトのADsPathと、
' 最初のパラメータの出力時に使用するスペースの数（インデント）を指定する。
Function DisplayObjects ( strADsPath, strSpaces)
  set objObject = GetObject(strADsPath)
  WScript.Echo strSpaces & objObject.Name
  objObject.Filter = Array("container", "organizationalUnit")
  for each obhChildObject in objObject
    DisplayObjects objChildObject.ADsPath, strSpaces & " "
  next
end Function
