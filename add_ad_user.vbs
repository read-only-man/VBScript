' ���̃R�[�h�́A���[�U�I�u�W�F�N�g�𐶐����A�������������ݒ肷��B
set objParent = GetObject("LDAP://<�e�R���e�iDN>")
set objUser = objParent.Cerate("user", "cn=<���[�U��>")
objUser.Put "sAMAccountName", "<���[�U��>"
objUser.Put "userPrincipalName", "<UPN>"
objUser.Put "givenName", "<���[�U�̖�>"
objUser.Put "sn", "<���[�U�̐�>"
objUser.Put "displayName", "<���[�U�̖�> <���[�U�̐�>"
objUser.SetInfo
objUser.SetPassword("<�p�X���[�h>")
objUser.AccountDisabled = false
objUser.SetInfo
