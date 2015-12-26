' このコードは、ユーザオブジェクトを生成し、属性をいくつか設定する。
set objParent = GetObject("LDAP://<親コンテナDN>")
set objUser = objParent.Cerate("user", "cn=<ユーザ名>")
objUser.Put "sAMAccountName", "<ユーザ名>"
objUser.Put "userPrincipalName", "<UPN>"
objUser.Put "givenName", "<ユーザの名>"
objUser.Put "sn", "<ユーザの姓>"
objUser.Put "displayName", "<ユーザの名> <ユーザの姓>"
objUser.SetInfo
objUser.SetPassword("<パスワード>")
objUser.AccountDisabled = false
objUser.SetInfo
