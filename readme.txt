Dim objExcel, objWorkbook, objSheet, objFSO, objFolder, objFile
Dim row, folderName, baseDirectory, fileContent

' Set the path of your Excel file
excelFilePath = "C:\path\to\your\excel.xlsx" ' Change this
baseDirectory = "C:\path\to\destination" ' Change this

' File content for index.test.tsx
fileContent = "import React from 'react';" & vbCrLf & vbCrLf & _
              "test('Sample test case', () => {" & vbCrLf & _
              "  expect(true).toBe(true);" & vbCrLf & _
              "});"

' Create File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Create Excel Application
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False ' Keep Excel hidden
Set objWorkbook = objExcel.Workbooks.Open(excelFilePath)
Set objSheet = objWorkbook.Sheets(1) ' First sheet

row = 2 ' Assuming first row has headers, start from row 2

' Loop through rows in column A (first column)
Do While objSheet.Cells(row, 1).Value <> ""
    folderName = Trim(objSheet.Cells(row, 1).Value)
    If folderName <> "" Then
        folderPath = baseDirectory & "\" & folderName
        ' Create folder if it doesn't exist
        If Not objFSO.FolderExists(folderPath) Then
            objFSO.CreateFolder folderPath
        End If
        ' Create index.test.tsx file inside the folder
        Set objFile = objFSO.CreateTextFile(folderPath & "\index.test.tsx", True)
        objFile.Write fileContent
        objFile.Close
    End If
    row = row + 1
Loop

' Cleanup
objWorkbook.Close False
objExcel.Quit

Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
Set objFSO = Nothing

MsgBox "Folders and test files created successfully!", vbInformation, "Done"