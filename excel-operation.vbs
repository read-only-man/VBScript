' Excel����̃T���v��

On Error Resume Next
set objExcel = CreateObject("Excel.Application")
if Err.Number <> 0 then
  Wscript.Echo "Excel Application no installed."
  Wscript.Quit -1
end if

' �Ȍ�̃G���[�͕֋X��A��������B
on Error GOTO 0

' �V�������[�N�u�b�N���쐬����B
objExcel.Workbooks.Add

' ���[�N�V�[�g�Ƀo�C���h����B
set objSheet = objExcel.ActiveWorkbook.WorkSheets(1)
objSheet.Name = "Processes"

' �X�v���b�h�V�[�g�̑���@���̕ӂ�Excel �}�N���Ƃ�����̂Ŋ���
objSheet.Cells(1,1).Value = "HogeHoge"

' �����Ƃ�
objExcel.Columns(1).ColumnWidth = 20

' �ۑ��ƏI��
objExcel.ActiveWorkbook.SaveAs "C:\work\test.xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
