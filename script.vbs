Dim args, objExcel

Set args=wscript.Arguments
Set objExcel=CreateObject("Excel.Application")

objExcel.workbooks.Open args(0)
objExcel.visible = True

objExcel.Run "BirthdayFamily"
objExcel.Run "BirthdayFriends"

objExcel.Activeworkbook.Save
objExcel.Activeworkbook.Close(0)
objExcel.Quit