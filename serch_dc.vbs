' ���̃R�[�h�́A�h���C���ɂ����āA���̃X�N���v�g�����s����Ă���R���s���[�^����
' �ł��߂��h���C���R���g���[������������B
' ----- START CONFIGRATION -----
strDomain = "<�h���C����DNS��>"
' -----  END  CONFIGRATION -----
set objIadsTools = CreateObject("IADsTools.DCFunctions")
objIadsTools.DsGetDcName(Cstr(strDomain))
WScript.Echo "DC: " & objIadsTools.DcName
WScript.Echo "DC Site: " & objIadsTools.DcSiteName
WScript.Echo "Client Site: " & objIadsTools.ClientSiteName
