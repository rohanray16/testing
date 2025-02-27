Dim objExcel, objWorkbook, objSheet, objFSO, objFolder
Dim excelFilePath, targetDirectory, lastRow, folderName, folderPath
Const xlUp = -4162

' Define the Excel file path (update as per your file location)
excelFilePath = "C:\Path\To\Your\File.xlsx"

' Define the directory where folders should be deleted
targetDirectory = "C:\Path\To\Target\Directory"

' Create Excel application object
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False ' Keep Excel hidden

' Open the workbook
Set objWorkbook = objExcel.Workbooks.Open(excelFilePath)
Set objSheet = objWorkbook.Sheets(1) ' Read from the first sheet

' Find the last row with data in column 1 (A)
lastRow = objSheet.Cells(objSheet.Rows.Count, 1).End(xlUp).Row

' Create File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Loop through each row in column A
For i = 1 To lastRow
    folderName = Trim(objSheet.Cells(i, 1).Value)
    
    ' Skip empty values
    If folderName <> "" Then
        folderPath = targetDirectory & "\" & folderName
        
        ' Check if the folder exists and delete it
        If objFSO.FolderExists(folderPath) Then
            objFSO.DeleteFolder folderPath, True
            WScript.Echo "Deleted: " & folderPath
        Else
            WScript.Echo "Folder not found: " & folderPath
        End If
    End If
Next

' Cleanup
objWorkbook.Close False
objExcel.Quit
Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
Set objFSO = Nothing

WScript.Echo "Process completed."