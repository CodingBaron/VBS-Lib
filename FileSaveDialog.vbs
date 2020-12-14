Option Explicit


Function SaveDlg(fileName,fileType)
    Set objDialog = CreateObject( "SAFRCFileDlg.FileSave" )
    ' Note: If no path is specified, the "current" directory will
    '       be the one remembered from the last "SAFRCFileDlg.FileOpen"
    '       or "SAFRCFileDlg.FileSave" dialog!
    objDialog.FileName = fileName
    ' Note: The FileType property is cosmetic only, it doesn't
    '       automatically append the right file extension!
    '       So make sure you type the extension yourself!
    objDialog.FileType = fileType
    If objDialog.OpenFileSaveDlg Then
        WScript.Echo "objDialog.FileName = " & objDialog.FileName
    End If
End Function