VERSION 5.00
Begin VB.UserControl ReadOutput 
   BackColor       =   &H80000007&
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   975
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   765
   ScaleWidth      =   975
   Begin VB.Image imgIcon 
      Height          =   735
      Left            =   0
      Picture         =   "ReadOutput.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "ReadOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 2.0 OF READ OUTPUT [CONTROL]

'You may use this code in your project as long as you dont claim its yours ;)

'This program reads the output of CLI (Command Line Interface) Applications.
'Examples of CLI Applications are:
'   -PING.EXE
'   -NETSTAT
'   -TRACERT
'This program will grab the output and call events so that you can process the commands.
'NOTE:  I got about 50% of this code from some site about 2 years ago, just found it and fixed some bugs
'       and transformed it into a user-friendly control.
'
'Please vote if you like, complaint about bugs if you find any, but I also appreciate comments ;)
'Thanks again
'-Endra


'ADDONS:
'   -The Terminate Process ID was originaly coded by Nick Campbeln. Thanks a bunch ;)
        '-His code is surrounded by POUND (##) signs.

Option Explicit 'force declarations of variables


'#####################################################################################
'#####################################################################################
'#####################################################################################
    '#### Functions/Consts used for CloseProcess()
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Private Const WM_CLOSE As Long = &H10
Private Const WM_DESTROY As Long = &H2
'Private Const WM_QUERYENDSESSION = &H11
Private Const WM_ENDSESSION = &H16
Private Const PROCESS_TERMINATE As Long = &H1

    '#### Functions/Consts/Types used for GetVersionEx() API
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long           '#### NT: Build Number, 9x: High-Order has Major/Minor ver, Low-Order has build
    PlatformID As Long
    szCSDVersion As String * 128    '#### NT: ie- "Service Pack 3", 9x: 'arbitrary additional information'
End Type
'Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

    '#### Functions/Consts used for CloseAll()
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

    '#### Required local vars
Private g_bIsInit As Boolean
Private g_bIs9x As Boolean
'#####################################################################################
'#####################################################################################
'#####################################################################################


'private variables
Dim sCommand As String
Dim bProcessing As Boolean
Dim bCancelProcess As Boolean

'Public Events
Public Event Error(ByVal Error As String, LastDLLError As Long) 'Errors go here
Public Event GotChunk(ByVal sChunk As String, ByVal LastChunk As Boolean)                   'Chunk Output detected, launch this event
Public Event Complete()                                         'Raise event when its done reading output
Public Event Starting()                                         'Raised when you can start the reading
Public Event Canceled()                                         'Raised when you canceled.
'The following are all API Calls and Types.
'You could probably find more information on them if you google them so I wont comment them at all.
Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long
      
Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long
      
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
      
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, _
    lpProcessAttributes As Any, lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As Any, lpProcessInformation As Any) As Long
      
Private Declare Function CloseHandle2 Lib "kernel32" (ByVal hObject As Long) As Long

'The following are simply constants that dont need changing during the program.
'DO NOT EDIT THESE!

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0

Private Sub UserControl_Initialize()
    On Error Resume Next
    'doesnt error out of stack space
    UserControl.Height = imgIcon.Height
    UserControl.Width = imgIcon.Width
    bProcessing = False
    bCancelProcess = False
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    'doesnt error out of stack space
    UserControl.Height = imgIcon.Height
    UserControl.Width = imgIcon.Width
End Sub

'The following function executes the command line and returns the output via events
Private Function ExecuteApp(sCmdline As String) As String
    Dim proc As PROCESS_INFORMATION, Ret As Long
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long 'The handle used to read from the pipe.
    Dim hWritePipe As Long 'The pipe where StdOutput and StdErr will be redirected to.
    Dim sOutput As String
    Dim lngBytesRead As Long, sBuffer As String * 256
    bProcessing = True
    sa.nLength = Len(sa)
    sa.bInheritHandle = True
      
    Ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    If Ret = 0 Then
        bProcessing = False
        RaiseEvent Error("CreatePipe failed.", Err.LastDLLError)
        Exit Function
    End If

    start.cb = Len(start)
    start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    ' Redirect the standard output and standard error to the same pipe
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
    start.wShowWindow = SW_HIDE
       
    ' Start the shelled application:
    ' if you program has to work only on NT you don't need the "conspawn "
    'ret = CreateProcessA(0&, "conspawn " & sCmdline, sa, sa, True, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    Ret = CreateProcessA(0&, Environ("ComSpec") & " /c " & sCmdline, sa, sa, True, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    If Ret = 0 Then
        bProcessing = False
        RaiseEvent Error("CreateProcess failed.", Err.LastDLLError)
        Exit Function
    End If
   
    ' The handle wWritePipe has been inherited by the shelled application
    ' so we can close it now
    CloseHandle hWritePipe

    ' Read the characters that the shelled application
    ' has outputed 256 characters at a time
    RaiseEvent Starting
    bCancelProcess = False
    Do
        DoEvents
        If bCancelProcess = True Then
            Exit Do
        End If
        Ret = ReadFile(hReadPipe, sBuffer, 256, lngBytesRead, 0&)
         
        sOutput = Left$(sBuffer, lngBytesRead)
'        MsgBox sOutput
        'frmReadOutputExample.lstout.AddItem Replace(sOutput, vbCrLf, "")
        If Ret = 0 Then
            RaiseEvent GotChunk(sOutput, True)  'no more chunks to read
            RaiseEvent Complete
            Exit Do
        Else
            RaiseEvent GotChunk(sOutput, False) 'more chunks to read.
        End If
    Loop While Ret <> 0 ' if ret = 0 then there is no more characters to read
    If bCancelProcess = True Then
        If CloseProcess(proc.dwProcessId) = True Then
            RaiseEvent Canceled
        Else
            RaiseEvent Error("Cannot close process id: " & proc.dwProcessId, 1203)
        End If
    End If

    CloseHandle proc.hProcess
    CloseHandle proc.hThread
    CloseHandle hReadPipe
    bProcessing = False
    bCancelProcess = False
End Function

Public Property Let SetCommand(ByVal sCommandVal As String)
    sCommand = sCommandVal
End Property

Public Property Get SetCommand() As String
    SetCommand = sCommand
End Property

Public Sub CancelProcess()
    If bProcessing = True Then
        bCancelProcess = True
    Else
        RaiseEvent Error("Not currently processing a command!", 1202)
    End If
End Sub

Public Sub ProcessCommand()
    If Len(sCommand) = 0 Then
        RaiseEvent Error("Invalid Command.", 1200)
        Exit Sub
    End If
    If bProcessing = True Then
        RaiseEvent Error("Currently processing a command!", 1201)
        Exit Sub
    End If
    ExecuteApp """" & sCommand & """"
End Sub

'#####################################################################################
'#####################################################################################
'#####################################################################################
'#####################################################################################
'# Public Functions
'#####################################################################################
'#########################################################
'# Ends a process according to the passed eMode
'#########################################################
Public Function CloseProcess(ByVal lProcessID As Long, Optional ByVal uExitCode As Long = 0) As Boolean
    Dim lTemp As Long

        '#### If we have not yet been initilized, call InitCloseProcess()
    If (Not g_bIsInit) Then Call InitCloseProcess

        '#### If we're running under Win95 or Win98 (WinME seems to process the other method correctly)
    If (g_bIs9x) Then
            '#### If we successfully send the 'Windows is closing' message to the lProcessID
        If (CloseAll(lProcessID, WM_ENDSESSION, True)) Then
                '#### Since the window has accepted the 'Windows is closing' message, we can now safely terminate the process
                '#### Collect a process handle in lTemp for lProcessID
            lTemp = OpenProcess(PROCESS_TERMINATE, False, lProcessID)

                '#### If lTemp is invalid, return false
            If (lTemp = 0) Then
                CloseProcess = False

                '#### Else the collected process handle is valid
            Else
                    '#### TerminateProcess() returns non-zero (true) on success and zero (false) on failure
                CloseProcess = CBool(TerminateProcess(lTemp, uExitCode))

                    '#### Close the open lTemp
                Call CloseHandle2(lTemp)
            End If

            '#### Else we could not communicate with the process
        Else
            CloseProcess = False
        End If

        '#### Else we're under a system that correctly handles the WM_CLOSE message
    Else
        CloseProcess = CloseAll(lProcessID, WM_CLOSE)
    End If
End Function



'#####################################################################################
'# Private Functions
'#####################################################################################
'#########################################################
'# Initilizes the module variables
'#########################################################
Private Sub InitCloseProcess()
    Dim uOSInfo As OSVERSIONINFO

        '#### Setup the uOSInfo UDT to determine the value of g_bIsNT4
    With uOSInfo
        .OSVSize = Len(uOSInfo)
        .szCSDVersion = Space(128)

            '#### Get the OS info, setting g_bIs9x accordingly
        Call GetVersionEx(uOSInfo)
        g_bIs9x = (.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (.dwVerMajor > 4) Or (.dwVerMajor = 4 And .dwVerMinor > 0) Or _
         (.PlatformID = VER_PLATFORM_WIN32_WINDOWS And .dwVerMajor = 4 And .dwVerMinor = 0) 'Or _
'!' WinME         (.PlatformID = VER_PLATFORM_WIN32_WINDOWS And .dwVerMajor = 4 And .dwVerMinor = 90)
    End With

        '#### Set g_bIsInit to true
    g_bIsInit = True
End Sub


'#########################################################
'# Posts the eMessage to all of the windows with the matching lProcessID
'#########################################################
Private Function CloseAll(ByVal lProcessID As Long, Optional ByVal eMessage As Long = WM_CLOSE, Optional ByVal wParam As Long = 0) As Boolean
    Dim hWndChild As Long
    Dim lThreadProcessID As Long

        '#### Get the Desktop handle while getting the first child under the Desktop and default the return value
    hWndChild = GetWindow(GetDesktopWindow(), GW_CHILD)
    CloseAll = False

        '#### While we still have hWndChild(en) to look at
    Do While (hWndChild <> 0)
            '#### If this is a parent window
        If (GetParent(hWndChild) = 0) Then
                '#### Get the lThreadProcessID of the window
            Call GetWindowThreadProcessId(hWndChild, lThreadProcessID)

                '#### If we have a match with the ProcessIDs
            If (lProcessID = lThreadProcessID) Then
                    '#### Post the message to the process and set the return value to true
                Call PostMessage(hWndChild, eMessage, wParam, 0&)
                CloseAll = True
            End If
        End If

            '#### Move onto the next hWndChild
        hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
    Loop
End Function
'#####################################################################################
'#####################################################################################
'#####################################################################################
'#####################################################################################
