' ���̃R�[�h�̓��[�J��DNS�L���b�V�����t���b�V������B
' �����ړI�Ƃ����X�N���v�e�B���O�C���^�t�F�C�X�͂Ȃ����߁A
' �V�F�����N������ipconfig /flushdns�R�}���h���Ăяo���Ă���B
strCommand = "ipconfig /flushdns"
run_command(strCommand)

' ���̃R�[�h�̓��[�J���z�X�g��DNS���R�[�h��o�^����
' �����ړI�Ƃ����X�N���v�e�B���O�C���^�t�F�C�X�͂Ȃ����߁A
' �V�F�����N������ipconfi /registerdns�R�}���h���Ăяo���Ă���B
strCommand = "ipconfig /registerdns"
run_command(strCommand)

function run_command(strCommand)
  set objWshShell = WScript.CreateObject("WScript.Shell")
  intRC= objWshShell.Run(strCommand, 0 ,TRUE)
  if intRC <> 0 then
    WScript.Echo "Error returned from running the command <" & strCommand & "> : " & intRC
  else
    WScript.Echo "Command excuted sucessfully <" & strCommand & ">"
  end if
end function
