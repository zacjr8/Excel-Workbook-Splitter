Sub GenerateSeparateWorkbooks()
    Dim ws As Worksheet
    Dim newWorkbook As Workbook
    Dim mainWorkbook As Workbook
    Dim sheetName As String
    Dim newWorkbookName As String
    Dim folderPath As String

    ' Set the main workbook
    Set mainWorkbook = ThisWorkbook

    ' Define the folder path where you want to save the new workbooks
    ' Change this to your desired folder path
    folderPath = "C:\YourFolderPath\" ' Update this path

    ' Loop through each worksheet, starting from the second one
    For Each ws In mainWorkbook.Worksheets
        If ws.Index > 1 Then ' Skip the first sheet
            ' Create a new workbook
            Set newWorkbook = Workbooks.Add
            
            ' Copy the current worksheet to the new workbook
            ws.Copy Before:=newWorkbook.Sheets(1)

            ' Delete the default sheet in the new workbook
            Application.DisplayAlerts = False
            newWorkbook.Sheets(2).Delete
            Application.DisplayAlerts = True

            ' Set the new workbook name
            sheetName = ws.Name
            newWorkbookName = "Document Repository - " & sheetName & ".xlsx"

            ' Save the new workbook
            newWorkbook.SaveAs folderPath & newWorkbookName, FileFormat:=xlOpenXMLWorkbook
            newWorkbook.Close SaveChanges:=False
        End If
    Next ws

    MsgBox "Workbooks generated successfully!", vbInformation
End Sub
