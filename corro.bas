Attribute VB_Name = "corro"
Option Compare Database
Option Explicit

Public Function getFolderPath() As String

'returns a directory

    Dim myPath As Office.FileDialog
    Dim folderPath As String

    Set myPath = Application.FileDialog(msoFileDialogFolderPicker)

    myPath.AllowMultiSelect = False
               
    If myPath.Show = True Then
        folderPath = myPath.SelectedItems(1)
        Else: MsgBox "No file selected!", vbOKOnly + vbExclamation
    End If

    Set myPath = Nothing
    
    If folderPath <> "" Then
        getFolderPath = folderPath
    End If
    
End Function

Public Function getFilePath(ByVal fileType As String) As Variant

' returns a file path or an array of file paths

    Dim myPath As Office.FileDialog
    Dim tPath As Variant
    Dim arrPaths() As String
    Dim fileCount As Integer
    Dim i As Long
    
    Set myPath = Application.FileDialog(msoFileDialogFilePicker)

        With myPath
            .AllowMultiSelect = True
            .Filters.Clear
            If fileType = "xls" Then
                .Filters.Add "Excel Files", "*.xls"
            ElseIf fileType = "xlsx" Then
                .Filters.Add "Excel Files", "*.xlsx"
            ElseIf fileType = "txt" Then
                .Filters.Add "Text Files", "*.txt"
            ElseIf fileType = "accdb" Then
                .Filters.Add "Access Files", "*.accdb, *.mdb"
            End If
        End With
               
        If myPath.Show = True Then
        
           fileCount = myPath.SelectedItems.Count
           
            ReDim arrPaths(0 To fileCount) As String
                For Each tPath In myPath.SelectedItems
                    arrPaths(i) = tPath
                    i = i + 1
                Next
            Else: MsgBox "No file selected!", vbOKOnly + vbExclamation
        End If

    Set myPath = Nothing
    
    If fileCount > 1 Then
        getFilePath = arrPaths
    ElseIf fileCount = 1 Then
        getFilePath = arrPaths(0)
    End If

End Function

Public Function MoveFiles(ByVal fromPath As String, toPath As String)

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If corro.fileExists(fromPath) Then
        fso.CopyFile fromPath, toPath
    Else: MsgBox "File not found"
    End If
    
    Set fso = Nothing

End Function

Public Sub Dialog(strmessage As String, showButton As Boolean)

    DoCmd.OpenForm "dialog", acNormal
    
    Forms!Dialog!ctlMessage.Caption = strmessage
    
    If showButton Then
        Forms!Dialog!BtnOK.Visible = True
    Else: Forms!Dialog!BtnOK.Visible = False
    End If

End Sub

Public Sub ClearTable(ByVal tableName As String)

    If tableExists(tableName) Then
        DoCmd.RunSQL ("DELETE FROM " & tableName)
    Else: MsgBox tableName & " not found to clear"
    End If

End Sub

Public Sub DropTable(ByVal tableName As String)

    DoCmd.RunSQL ("DROP TABLE " & tableName)

End Sub

Public Sub IndexTable(ByVal indexName As String, tableName As String, column As String)

    DoCmd.RunSQL ("CREATE INDEX " & indexName & " ON " & tableName & " (" & column & ")")

End Sub

Public Function removeNullRowsFromFile(filePath As String, minLength As Integer)

'adhoc ouput often adds trailing empty rows. this deletes those

    Dim fso As Object
    Dim file As Object
    Dim strLine As String
    Dim strNewContents As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(filePath, 1)

    Do Until file.AtEndOfStream
    
        strLine = file.Readline
        strLine = Trim(strLine)
    
        If Len(strLine) > minLength Then
            strNewContents = strNewContents & strLine & vbCrLf
        End If
    
    Loop

    file.Close
    Set file = fso.OpenTextFile(filePath, 2)
    file.Write strNewContents
    file.Close
    
    Set file = Nothing
    Set fso = Nothing

End Function

Public Function tableExists(strTableName As String) As Boolean
    
'check if a table exists within the current database
    
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Set db = CurrentDb
    On Error Resume Next
    Set td = db.TableDefs(strTableName)
    tableExists = (Err.Number = 0)
    Err.Clear
    
End Function
Public Function queryExists(strQueryName As String) As Boolean
    
'check if a table exists within the current database
    
    Dim db As DAO.Database
    Dim td As DAO.QueryDef
    Set db = CurrentDb
    On Error Resume Next
    Set td = db.TableDefs(strTableName)
    tableExists = (Err.Number = 0)
    Err.Clear
    
End Function

Public Function removeIdCert()

    Dim appPath As String
    Dim wsh As Object
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    
    Set wsh = CreateObject("WScript.Shell")
    appPath = Application.CurrentProject.path

    wsh.Run "cmd /c type " & appPath & "\SCRIPT\remove_id_cert.txt | powershell", windowStyle, waitOnReturn
    
    Set wsh = Nothing
    
End Function

Sub Test()

    MsgBox Application.CurrentProject.path

End Sub

Public Function fileExists(ByVal path As String) As Boolean

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileExists = fso.fileExists(path)
    Set fso = Nothing
    
End Function

Public Function parentFolder(ByVal path As String) As String

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    parentFolder = fso.parentFolder(path)
    
    Set fso = Nothing
    
End Function

Public Function fileName(ByVal path As String) As String

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fileName = fso.GetFileName(path)
    
    Set fso = Nothing

End Function


