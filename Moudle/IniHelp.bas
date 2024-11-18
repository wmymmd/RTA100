Attribute VB_Name = "IniHelp"
Option Explicit
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" _
(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Const ProcDict_Path As String = "\Config\ProcDict.ini"
Public Const Function_Path As String = "\Config\Function.ini"
Public Const DeviceInfo_Path As String = "\Config\DeviceInfo.ini"
Public Const ModbusRtu_Path As String = "\Config\ModbusRtu.ini"

Public Function Readini(Key As String)
     Dim result As Long
     Dim buffer As String
     buffer = String(255, 0)
     result = GetPrivateProfileString("Procdict", Key, "", buffer, Len(buffer), App.Path + ProcDict_Path)
        If result > 0 Then
              buffer = Left(buffer, result)
              Else
              buffer = ""
        End If
        Readini = buffer
End Function


Public Function CommnonReadini(Section As String, Key As String, FilePath As String)
     Dim result As Long
     Dim buffer As String
     buffer = String(255, 0)
     result = GetPrivateProfileString(Section, Key, "", buffer, Len(buffer), FilePath)
        If result > 0 Then
              buffer = Left(buffer, result)
              Else
              buffer = ""
        End If
        CommnonReadini = buffer
End Function


'Public Sub DeleteKeyFromIni(ByVal iniFileName As String, ByVal sectionName As String, ByVal keyName As String)
'    WritePrivateProfileString sectionName, keyName, "", iniFileName
'End Sub


Public Function GetKeysInSection(ByVal sectionName As String, ByVal iniFilePath As String) As String()
    Dim buffer As String * 1024
    Dim result As Long
    Dim Keys() As String
    
    result = GetPrivateProfileSection(sectionName, buffer, 1024, iniFilePath)
    If result > 0 Then
        Keys = Split(Trim(buffer), vbNullChar)
    End If
    GetKeysInSection = Keys
End Function


Public Function GetKeyCountInSection(ByVal strSection As String, ByVal strINIFile As String) As Long
    Dim strBuffer As String
    Dim lngBufferSize As Long
    Dim lngResult As Long
    Dim strTemp() As String
    lngBufferSize = 4294967296#
    strBuffer = String(lngBufferSize, 0)
    lngResult = GetPrivateProfileSection(strSection, strBuffer, lngBufferSize, strINIFile)
    If lngResult > 0 Then
        strBuffer = Left(strBuffer, lngResult)
        strTemp = Split(strBuffer, vbNullChar)
        GetKeyCountInSection = UBound(strTemp)
    Else
        GetKeyCountInSection = 0
    End If
End Function


Public Function SectionExistsInIni(ByVal iniFilePath As String, ByVal sectionName As String) As Boolean
    Dim buffer As String
    buffer = String(255, Chr(0))
    GetPrivateProfileString sectionName, vbNullString, "", buffer, Len(buffer), iniFilePath
    SectionExistsInIni = (Len(Trim(buffer)) > 0)
End Function

Public Function KeyExistsInIni(ByVal iniFilePath As String, ByVal sectionName As String, ByVal keyName As String) As Boolean
    Dim buffer As String
    buffer = String(255, Chr(0))
    GetPrivateProfileString sectionName, keyName, "", buffer, Len(buffer), iniFilePath
    KeyExistsInIni = (Len(Trim(buffer)) > 0)
End Function

