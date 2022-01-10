Attribute VB_Name = "FuncMe"
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" _
    (ByVal pszPath As String) As Long
    Private Declare Function GetModuleFileNameW Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpFilename As Long, ByVal nSize As Long) As Long
    Public Declare Function ShellExecute Lib "Shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function GetEXEName() As String
    Const MAX_PATH = 260&

    GetEXEName = Space$(MAX_PATH - 1&)
    GetEXEName = Left$(GetEXEName, GetModuleFileNameW(0&, StrPtr(GetEXEName), MAX_PATH))
    GetEXEName = Right$(GetEXEName, Len(GetEXEName) - InStrRev(GetEXEName, "\"))
End Function



