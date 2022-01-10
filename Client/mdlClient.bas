Attribute VB_Name = "mdlClient"
'Winsock Client/Server, created by the KPD-Team 2001
'This file can be downloaded from http://www.allapi.net/
'For questions or comments, contact us at KPDTeam@AllAPI.net
'
' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.

Public Const SERVER_PORT As Long = 1019
Public Const GWL_WNDPROC = (-4)
Public Type HostEnt
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type
Public Declare Sub CopyMemoryIP Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public theSocket As Long
Private PrevProc As Long
'Function to retrieve the IP address
Public Function GetIPAddress() As String
    Dim sHostName As String * 256
    Dim lpHost As Long
    Dim Host As HostEnt
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim I As Integer
    Dim sIPAddr As String
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPAddress = ""
        'MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
        Exit Function
    End If
    sHostName = Trim$(sHostName)
    lpHost = gethostbyname(sHostName)
    If lpHost = 0 Then
        GetIPAddress = ""
        'MsgBox "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
        Exit Function
    End If
    CopyMemoryIP Host, lpHost, Len(Host)
    CopyMemoryIP dwIPAddr, Host.hAddrList, 4
    ReDim tmpIPAddr(1 To Host.hLen)
    CopyMemoryIP tmpIPAddr(1), dwIPAddr, Host.hLen
    For I = 1 To Host.hLen
        sIPAddr = sIPAddr & tmpIPAddr(I) & "."
    Next
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
End Function
'For more information about subclassing,
'visit our subclassing tutorial on http://www.allapi.net/
Public Sub StartSubclass(F As Form)
    PrevProc = SetWindowLong(F.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub StopSubclass(F As Form)
    If PrevProc <> 0 Then SetWindowLong F.hwnd, GWL_WNDPROC, PrevProc
End Sub
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WINSOCK_MESSAGE Then
        ProcessMessage wParam, lParam
    Else
        WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
    End If
End Function
'Process socket messages
Public Function ProcessMessage(ByVal wParam As Long, ByVal lParam As Long) 'wParam = Socket Handle, lParam = connection message
    Select Case lParam
        Case FD_ACCEPT
        Case FD_CONNECT
            'MsgBox "Connect"
        Case FD_WRITE
        Case FD_READ
            Dim sTemp As String, lRet As Long, szBuf As String
            Do
                szBuf = String(256, 0)
                lRet = recv(wParam, ByVal szBuf, Len(szBuf), 0)
                If lRet > 0 Then sTemp = sTemp + Left$(szBuf, lRet)
            Loop Until lRet <= 0
            If LenB(sTemp) > 0 Then AddText sTemp
        Case FD_CLOSE
            frmMain.lblstat.Caption = "Lost connection with server..."
    End Select
End Function
Public Sub AddText(sInput As String)
    frmMain.txtReceived.Text = sInput '+ IIf(Right$(sInput, 2) = vbCrLf, "", vbCrLf)
    'frmMain.txtReceived.SelStart = Len(frmMain.txtReceived.Text)
    
End Sub
Public Function IsInIDE() As Boolean
On Local Error GoTo ErrHandler
    Debug.Print 1 / 0
Exit Function
ErrHandler:
    IsInIDE = True
End Function
