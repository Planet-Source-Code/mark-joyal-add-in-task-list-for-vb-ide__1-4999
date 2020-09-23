Attribute VB_Name = "modGlobals"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Global gblFile As String
Global gVBInstance  As VBIDE.VBE
Global gwinWindow   As VBIDE.Window
Global gblMouseClick As Integer

Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function

Function WriteToIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    WriteToIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

Public Function FileExists(strFile As String) As String
    On Error Resume Next 'Doesn't raise error - FileExists will be False
    FileExists = Dir(strFile, vbHidden) <> ""
End Function
