Sub getFilenameImages()
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim i As Long
Dim source As String

source = Cells(2, 4).Value
'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder(source)
i = 7
'loops through each file in the directory and prints their names and path
For Each objFile In objFolder.Files
    'print file name
    Cells(i, 2) = objFile.Name
    'print file path
    'Cells(i + 1, 3) = objFile.Path
    i = i + 1
Next objFile

i = 7
For Each objFile In objFolder.SubFolders
    'print file name
    Cells(i, 3) = objFile.Name
    'print file path
    'Cells(i + 1, 3) = objFile.Path
    i = i + 1
Next objFile
End Sub
