# Excel-VBA-folder-files
Excel-VBA-folder-files

```vba
Sub moveFiles()

Dim oFSO As Object 'FileSystemObject
Dim oFolder As Object 'Scan target folder
Dim oFile As Object 'Found target file
Dim fileObj As Object 'File object for mod date
Dim fileModDate As String
Dim fileModDateEpoch As Double 'Fake windows epoch
Dim fileDate As Integer 'Strip after decimal and preserve date as integer
Dim shift As String
Dim osPathSeperator As String 'Operation system path seperator
Dim mainFolder As String
Dim newFolder As String
Dim mainFolderPath As String
Dim newFolderPath As String

osPathSeperator = "\" 'Define for windows
mainFolder = "Fotolar"
newFolder = "Yeni"
mainFolderPath = ActiveWorkbook.Path & osPathSeperator & mainFolder
newFolderPath = ActiveWorkbook.Path & osPathSeperator & newFolder

'Define active path
'Debug.Print "Main Folder Path --> " & mainFolderPath
'Debug.Print "New Folder Path --> " & newFolderPath


'Create FileSystemObject
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Create working folders
If Not oFSO.FolderExists(newFolderPath) Then
  oFSO.CreateFolder newFolderPath
End If

If Not oFSO.FolderExists(mainFolderPath) Then
  oFSO.CreateFolder mainFolderPath
End If

'Scan dir
Set oFolder = oFSO.GetFolder(newFolderPath)


'Loop each file
For Each oFile In oFolder.Files
    
    Set fileObj = oFSO.GetFile(oFile) 'Get file object
    'fileModDate = fileObj.DateLastModified 'Read mod timestamp as string
    fileModDateEpoch = CDbl(fileObj.DateLastModified) 'Convert to epoch before converting string
    
    'Find shift
    If fileModDateEpoch >= Int(fileModDateEpoch) + (1 / 24 * 0) And fileModDateEpoch < Int(fileModDateEpoch) + (1 / 24 * 8) Then
      shift = Format(fileModDateEpoch, "yyyy-mm-dd") & " 00-08"
    ElseIf fileModDateEpoch >= Int(fileModDateEpoch) + (1 / 24 * 8) And fileModDateEpoch < Int(fileModDateEpoch) + (1 / 24 * 16) Then
      shift = Format(fileModDateEpoch, "yyyy-mm-dd") & " 08-16"
    ElseIf fileModDateEpoch >= Int(fileModDateEpoch) + (1 / 24 * 16) And fileModDateEpoch < Int(fileModDateEpoch) + (1 / 24 * 24) Then
      shift = Format(fileModDateEpoch, "yyyy-mm-dd") & " 16-24"
    End If
    
    ' Show shift
    'Debug.Print mainFolderPath & osPathSeperator & shift
    
    'Check if folder exist
    If Not oFSO.FolderExists(mainFolderPath & osPathSeperator & shift) Then
      'Create folder
      oFSO.CreateFolder mainFolderPath & osPathSeperator & shift
    End If
    
    
    'Move file to related folder
    'MoveFile doesnt have overwrite option
    'oFSO.MoveFile Source:=oFile, Destination:=mainFolderPath & osPathSeperator & shift & osPathSeperator & oFile.Name
    
    'Copy paste alternative
    oFSO.CopyFile oFile, mainFolderPath & osPathSeperator & shift & osPathSeperator & oFile.Name, True
    oFSO.DeleteFile oFile
        
    'Debug.Print oFile
    'Debug.Print oFile.Name
Next oFile

'Clear memory
Set oFSO = Nothing
Set oFolder = Nothing

End Sub
```




