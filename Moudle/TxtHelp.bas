Attribute VB_Name = "TxtHelp"
Option Explicit
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const ProcessStep_Path As String = "\Config\ProcessStep.txt"


Public Function ReadTextFileToArray(FilePath As String) As Collection
    Dim fileNumber As Integer
    Dim StringCollection As New Collection
    Dim lineArray() As String
    Dim fileContent As String
    Dim i As Long
    fileNumber = FreeFile
    Open FilePath For Input As #fileNumber
    i = 0
    Do While Not EOF(fileNumber)
       Line Input #fileNumber, fileContent
       If fileContent <> "" Then
          StringCollection.Add fileContent
       End If
    Loop
    Close #fileNumber
  Set ReadTextFileToArray = StringCollection
End Function

Public Sub WriteLineToFile(FilePath As String, text As String)
    Dim fileNumber As Integer
    fileNumber = FreeFile()
    Open FilePath For Append As #fileNumber
    Print #fileNumber, text
    Close #fileNumber
End Sub


Public Sub WriteLog(logMessage As String)
    Open App.Path + "\\ErrorLog.txt" For Append As #1
    Print #1, Now & ": " & logMessage
    Close #1
End Sub

Public Sub HideFile(FileName As String)
If dir(FileName) <> "" Then
Call SetFileAttributes(FileName, FILE_ATTRIBUTE_HIDDEN)
End If
End Sub

Public Sub ShowFile(FileName As String)
If dir(FileName) <> "" Then
Call SetFileAttributes(FileName, FILE_ATTRIBUTE_NORMAL)
End If
End Sub
'-------------------------------------¥[±K-------------------------------------------
Public Function EncryptDecrypt(ByVal str As String, ByVal key As Integer) As String
    Dim i As Integer
    Dim result As String
    Dim charValue As Integer
    
    For i = 1 To Len(str)
        charValue = Asc(Mid(str, i, 1))
        charValue = charValue Xor key
        result = result & Chr(charValue)
    Next i
    
    EncryptDecrypt = result
End Function
'------------------------------------------------------------------------------------
Public Sub ModifyTextFile(FilePath As String, ColIndex As Long, newText As String)
Dim i As Long
Dim fileNumber As Integer
Dim TxtCol As Collection
fileNumber = FreeFile()
Set TxtCol = ReadTextFileToArray(FilePath)
TxtCol.Remove ColIndex
If newText <> "" Then
TxtCol.Add newText, CStr(ColIndex)
End If
Open FilePath For Output As #fileNumber
Print #fileNumber, ""
For i = 1 To TxtCol.Count
If TxtCol(i) <> "" Then
Print #fileNumber, TxtCol(i)
End If
Next i
Close #fileNumber
End Sub

