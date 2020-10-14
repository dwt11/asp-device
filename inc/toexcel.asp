'п╢EXCELнд╪Ч
<%Set xlApplication = CreateObject("Excel.application")
xlApplication.Visible = False
xlApplication.Workbooks.Add()
Set xlWorksheet1 = xlApplication.Worksheets(1)
xlWorksheet1.Cells(1,1).Value = "a"
xlWorksheet1.Cells(1,2).Value = "b"
Set xlWorksheet2 = xlApplication.Worksheets(2)
xlWorksheet2.Cells(1,1).Value = "c"
xlWorksheet2.Cells(1,2).Value = "d"

xlWorksheet.SaveAs "test.xls"
xlApplication.Quit 
Set xlWorksheet1 = Nothing
Set xlWorksheet2 = Nothing
Set xlApplication = Nothing

%>