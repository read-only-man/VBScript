' ���[�����M�p���W���[��
' ���̃X�N���v�g�����s����ɂ�CDO�iCollaboration Data Objects�j���N���C�A���g�R���s���[�^�ɃC���X�g�[������Ă���K�v������B
' ���̂��߂ɂ� Outlook���C���X�g�[�����邩�AMicroSoft�̃T�C�g����CDO���ʂɃC���X�g�[������K�v������B

' ----- START CONFIGRATION -----
strFrom ="test@mail.com"
strTo   = "someone@mail.com"
strSub  = "mail title"
strBody = "this is mail test" & VBCRLF & "Hello"
strSMTPServer = "smtp.server.com"
' ----- END CONFIGRATION -------

set objEmail = CreateObject"CDO.Message")
objEmail.From = strFrom
objEmail.To = strTo
objEmail.Subject = strSub
objEmail.TextBody = strBody
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer
objEmail.Configuration.Fields.Update
objEmail.Send
WScript.Echo "Email sent"

objEmail = Nothing
