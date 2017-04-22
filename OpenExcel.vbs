
Dim objExcel

Set objExcel = CreateObject("Excel.Applicatiion")

objExcel.workbooks.Add

objExcel.ActiveWorkbook.SaveAs "C:\Users\Basit\Documents\Quality Assurance"
objExcel.Quit

Set objExcel = Nothing