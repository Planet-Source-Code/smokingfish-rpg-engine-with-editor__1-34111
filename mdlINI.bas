Attribute VB_Name = "mdlINI"
Declare Function WritePrivateProfileSection Lib _
"kernel32" Alias "WritePrivateProfileSectionA" _
(ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib _
"kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileSection Lib _
"kernel32" Alias "GetPrivateProfileSectionA" _
(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileString Lib _
"kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function WriteINI(ByVal lpAppName As String, ByVal IniKey As String, ByVal IniVal As String, ByVal lpFileName As String)
Dim lonStatus As Long
    lonStatus = WritePrivateProfileString(lpAppName, IniKey, IniVal, lpFileName)

End Function

