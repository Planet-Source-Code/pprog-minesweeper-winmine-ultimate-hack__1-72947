Attribute VB_Name = "Module1"
Option Base 1
''''''''
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
''''''
Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Public Const TH32CS_SNAPPROCESS As Long = 2&
Public WinmineHandle As Long

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Public Function GetWinmineHandle() As Boolean

Dim myProcess As PROCESSENTRY32
Dim mySnapshot As Long
Dim s As String
Dim i As Integer
Dim pid As Long

GetWinmineHandle = False
WinminePid = -1

myProcess.dwSize = Len(myProcess)
mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
ProcessFirst mySnapshot, myProcess

While ProcessNext(mySnapshot, myProcess)
    s = Trim$(myProcess.szexeFile)
    i = InStr(1, s, Chr(0))
    s = Mid$(s, 1, i - 1)
  If LCase(s) = "winmine.exe" Then
    pid = myProcess.th32ProcessID
        If pid = -1 Then Exit Function
        WinmineHandle = OpenProcess(&H1F0FFF, False, pid)
        GetWinmineHandle = True
        Exit Function
  End If
Wend

End Function

Public Function WriteMemoryLong(Address As Long, value As Long)

    If WinmineHandle = 0 Then Call GetWinmineHandle: Exit Function
    
    WriteProcessMemory WinmineHandle, Address, value, 4, 0&

End Function

Public Function WriteMemoryByte(Address As Long, value As Byte)

    If WinmineHandle = 0 Then Call GetWinmineHandle: Exit Function
    
    WriteProcessMemory WinmineHandle, Address, value, 1, 0&

End Function

Public Function ReadMemory(Address As Long) As Long

    Dim value As Long
    
    If WinmineHandle = 0 Then Call GetWinmineHandle: Exit Function
    
    ReadProcessMem WinmineHandle, Address, value, 1, 0&
    ReadMemory = value

End Function
