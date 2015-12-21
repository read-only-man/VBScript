' Excel操作のサンプル

On Error Resume Next
set objExcel = CreateObject("Excel.Application")
if Err.Number <> 0 then
  Wscript.Echo "Excel Application no installed."
  Wscript.Quit -1
end if

' 以後のエラーは便宜上、無視する。
on Error GOTO 0

' 新しいワークブックを作成する。
objExcel.Workbooks.Add

' ワークシートにバインドする。
set objSheet = objExcel.ActiveWorkbook.WorkSheets(1)
objSheet.Name = "Processes"

' スプレッドシートの操作　この辺はExcel マクロとかわらんので割愛
objSheet.Cells(1,1).Value = "HogeHoge"

' 書式とか
objExcel.Columns(1).ColumnWidth = 20

' 保存と終了
objExcel.ActiveWorkbook.SaveAs "C:\work\test.xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
